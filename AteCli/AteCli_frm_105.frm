VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   2055
   ClientTop       =   1725
   ClientWidth     =   11700
   Icon            =   "AteCli_frm_105.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7965
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   14049
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
         Height          =   5865
         Left            =   30
         TabIndex        =   33
         Top             =   1230
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   10345
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
         Begin VB.CommandButton cmd_DirCas 
            Caption         =   "="
            Height          =   315
            Left            =   1530
            TabIndex        =   2
            ToolTipText     =   "Obtener Direcci�n de Domicilio"
            Top             =   390
            Width           =   435
         End
         Begin VB.ComboBox cmb_ConLoc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   3510
            Width           =   765
         End
         Begin VB.ComboBox cmb_NomCar 
            Height          =   315
            Left            =   2010
            TabIndex        =   28
            Text            =   "cmb_Dep_NomCar"
            Top             =   5490
            Width           =   3315
         End
         Begin VB.TextBox txt_NomCar 
            Height          =   315
            Left            =   8220
            MaxLength       =   250
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   5490
            Width           =   3315
         End
         Begin VB.TextBox txt_Telef1 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   2040
            Width           =   1640
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   11
            Text            =   "cmb_DstDir"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   9
            Text            =   "cmb_DptDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   2010
            MaxLength       =   120
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   8220
            MaxLength       =   11
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   60
            Width           =   2355
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_NumFax 
            Height          =   315
            Left            =   8220
            MaxLength       =   12
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   2040
            Width           =   1640
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8220
            MaxLength       =   250
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8220
            TabIndex        =   10
            Text            =   "cmb_PrvDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8220
            MaxLength       =   120
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_IntDpt 
            Height          =   315
            Left            =   9870
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   720
            Width           =   1665
         End
         Begin VB.TextBox txt_NumVia 
            Height          =   315
            Left            =   8220
            MaxLength       =   15
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   720
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_Telef2 
            Height          =   315
            Left            =   3660
            MaxLength       =   12
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   2040
            Width           =   1640
         End
         Begin VB.ComboBox cmb_CodCiu 
            Height          =   315
            Left            =   2010
            TabIndex        =   16
            Text            =   "cmb_DptDir"
            Top             =   2370
            Width           =   9525
         End
         Begin VB.CommandButton cmd_BusEmp_Emp 
            Caption         =   "..."
            Height          =   315
            Left            =   10620
            TabIndex        =   22
            ToolTipText     =   "Obtener Direcci�n de Domicilio"
            Top             =   3840
            Width           =   435
         End
         Begin VB.TextBox txt_NumDoc_Emp 
            Height          =   315
            Left            =   8220
            MaxLength       =   11
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   3840
            Width           =   2355
         End
         Begin VB.ComboBox cmb_TipDoc_Emp 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   3840
            Width           =   3315
         End
         Begin VB.TextBox txt_RazSoc_Emp 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   4170
            Width           =   9525
         End
         Begin VB.TextBox txt_NomCom_Emp 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   4500
            Width           =   9525
         End
         Begin VB.TextBox txt_Telef1_Emp 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   4830
            Width           =   1640
         End
         Begin VB.TextBox txt_Telef2_Emp 
            Height          =   315
            Left            =   3660
            MaxLength       =   12
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   4830
            Width           =   1640
         End
         Begin EditLib.fpDoubleSingle ipp_IngNet 
            Height          =   315
            Left            =   2010
            TabIndex        =   17
            Top             =   2850
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
         Begin EditLib.fpDateTime ipp_IniAct 
            Height          =   315
            Left            =   2010
            TabIndex        =   18
            Top             =   3180
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
            TabIndex        =   34
            Top             =   2730
            Width           =   11475
            _Version        =   65536
            _ExtentX        =   20241
            _ExtentY        =   106
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
         End
         Begin Threed.SSPanel pnl_FlgEmp 
            Height          =   315
            Left            =   11100
            TabIndex        =   35
            Top             =   3840
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "NR"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
         End
         Begin EditLib.fpDateTime ipp_FecIng_Emp 
            Height          =   315
            Left            =   2010
            TabIndex        =   27
            Top             =   5160
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
         Begin VB.Label Label11 
            Caption         =   "Contrato Locaci�n:"
            Height          =   285
            Left            =   90
            TabIndex        =   66
            Top             =   3510
            Width           =   1785
         End
         Begin VB.Label lbl_General 
            Caption         =   "Ingreso Declarado (S/.):"
            Height          =   285
            Index           =   61
            Left            =   90
            TabIndex        =   59
            Top             =   2850
            Width           =   1755
         End
         Begin VB.Label lbl_General 
            Caption         =   "Cargo:"
            Height          =   285
            Index           =   62
            Left            =   90
            TabIndex        =   58
            Top             =   5490
            Width           =   975
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha Inicio Actividades:"
            Height          =   315
            Index           =   58
            Left            =   90
            TabIndex        =   57
            Top             =   3180
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Cargo (Especificar):"
            Height          =   285
            Index           =   57
            Left            =   6210
            TabIndex        =   56
            Top             =   5490
            Width           =   1665
         End
         Begin VB.Label lbl_General 
            Caption         =   "CIIU:"
            Height          =   285
            Index           =   39
            Left            =   90
            TabIndex        =   55
            Top             =   2370
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "N�mero Docum. Ident.:"
            Height          =   285
            Index           =   48
            Left            =   6210
            TabIndex        =   54
            Top             =   60
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   36
            Left            =   90
            TabIndex        =   53
            Top             =   60
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tel�fono (s):"
            Height          =   285
            Index           =   46
            Left            =   90
            TabIndex        =   52
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Distrito:"
            Height          =   315
            Index           =   45
            Left            =   90
            TabIndex        =   51
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Departamento:"
            Height          =   315
            Index           =   44
            Left            =   90
            TabIndex        =   50
            Top             =   1380
            Width           =   1425
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Index           =   43
            Left            =   90
            TabIndex        =   49
            Top             =   1050
            Width           =   1455
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre V�a:"
            Height          =   285
            Index           =   42
            Left            =   90
            TabIndex        =   48
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fax:"
            Height          =   285
            Index           =   55
            Left            =   6210
            TabIndex        =   47
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Referencia:"
            Height          =   285
            Index           =   54
            Left            =   6210
            TabIndex        =   46
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Provincia:"
            Height          =   315
            Index           =   53
            Left            =   6210
            TabIndex        =   45
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Index           =   52
            Left            =   6210
            TabIndex        =   44
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Index           =   51
            Left            =   6210
            TabIndex        =   43
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de V�a:"
            Height          =   285
            Index           =   41
            Left            =   90
            TabIndex        =   42
            Top             =   390
            Width           =   1545
         End
         Begin VB.Label lbl_General 
            Caption         =   "Raz�n Social:"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   41
            Top             =   4170
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "N�mero Docum. Ident.:"
            Height          =   285
            Index           =   2
            Left            =   6210
            TabIndex        =   40
            Top             =   3840
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   39
            Top             =   3840
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Index           =   4
            Left            =   90
            TabIndex        =   38
            Top             =   4500
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tel�fono (s):"
            Height          =   285
            Index           =   5
            Left            =   90
            TabIndex        =   37
            Top             =   4830
            Width           =   1815
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Index           =   6
            Left            =   90
            TabIndex        =   36
            Top             =   5160
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   60
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
            TabIndex        =   61
            Top             =   60
            Width           =   10125
            _Version        =   65536
            _ExtentX        =   17859
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes - Actividades Econ�micas - Profesional Independiente"
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
            Picture         =   "AteCli_frm_105.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   62
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
            TabIndex        =   63
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
         Begin VB.Label Label1 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   64
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   65
         Top             =   7140
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
            Picture         =   "AteCli_frm_105.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Simulaci�n de Cr�ditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10890
            Picture         =   "AteCli_frm_105.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10200
            Picture         =   "AteCli_frm_105.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_05"
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
         Call gs_SetFocus(ipp_IngNet)
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
      
      Call gs_SetFocus(ipp_IngNet)
   End If
End Sub

Private Sub cmb_ConLoc_Click()
   If cmb_ConLoc.ListIndex > -1 Then
      If cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex) = 1 Then
         cmb_TipDoc_Emp.Enabled = True
         txt_NumDoc_Emp.Enabled = True
         cmd_BusEmp_Emp.Enabled = True
         ipp_FecIng_Emp.Enabled = True
         cmb_NomCar.Enabled = True
         
         Call gs_SetFocus(cmb_TipDoc_Emp)
      Else
         cmb_TipDoc_Emp.ListIndex = -1
         txt_NumDoc_Emp.Text = ""
         pnl_FlgEmp.Caption = ""
         pnl_FlgEmp.Tag = ""
         txt_RazSoc_Emp.Text = ""
         txt_NomCom_Emp.Text = ""
         txt_Telef1_Emp.Text = ""
         txt_Telef2_Emp.Text = ""
         ipp_FecIng_Emp.Text = Format(Date, "dd/mm/yyyy")
         cmb_NomCar.ListIndex = -1
         txt_NomCar.Text = ""
         txt_NomCar.Enabled = False
         
         cmb_TipDoc_Emp.Enabled = False
         txt_NumDoc_Emp.Enabled = False
         cmd_BusEmp_Emp.Enabled = False
         txt_RazSoc_Emp.Enabled = False
         txt_NomCom_Emp.Enabled = False
         txt_Telef1_Emp.Enabled = False
         txt_Telef2_Emp.Enabled = False
         ipp_FecIng_Emp.Enabled = False
         cmb_NomCar.Enabled = False
         
         Call gs_SetFocus(cmd_Grabar)
      End If
   Else
      cmb_TipDoc_Emp.ListIndex = -1
      txt_NumDoc_Emp.Text = ""
      pnl_FlgEmp.Caption = ""
      pnl_FlgEmp.Tag = ""
      txt_RazSoc_Emp.Text = ""
      txt_NomCom_Emp.Text = ""
      txt_Telef1_Emp.Text = ""
      txt_Telef2_Emp.Text = ""
      ipp_FecIng_Emp.Text = Format(Date, "dd/mm/yyyy")
      cmb_NomCar.ListIndex = -1
      txt_NomCar.Text = ""
      txt_NomCar.Enabled = False
      
      cmb_TipDoc_Emp.Enabled = False
      txt_NumDoc_Emp.Enabled = False
      cmd_BusEmp_Emp.Enabled = False
      txt_RazSoc_Emp.Enabled = False
      txt_NomCom_Emp.Enabled = False
      txt_Telef1_Emp.Enabled = False
      txt_Telef2_Emp.Enabled = False
      ipp_FecIng_Emp.Enabled = False
      cmb_NomCar.Enabled = False
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
            Call gs_SetFocus(cmd_Grabar)
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
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub

Private Sub cmb_TipDoc_Emp_Click()
   If cmb_TipDoc_Emp.ListIndex > -1 Then
      Select Case cmb_TipDoc_Emp.ItemData(cmb_TipDoc_Emp.ListIndex)
         Case 1:     txt_NumDoc_Emp.MaxLength = 8
         Case 7:     txt_NumDoc_Emp.MaxLength = 11
         Case Else:  txt_NumDoc_Emp.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc_Emp)
End Sub

Private Sub cmb_TipDoc_Emp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Emp_Click
   End If
End Sub

Private Sub cmd_BusEmp_Emp_Click()
   If cmb_TipDoc_Emp.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc_Emp)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc_Emp.Text)) = 0 Then
      MsgBox "Debe ingresar el N�mero de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc_Emp)
      Exit Sub
   End If
   
   If cmb_TipDoc.ItemData(cmb_TipDoc_Emp.ListIndex) = 7 Then
      If Len(Trim(txt_NumDoc_Emp.Text)) <> 11 Then
         MsgBox "El N�mero de Documento ingresado no es correcto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc_Emp)
         Exit Sub
      End If
   End If
   
   'Buscando Empresa
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(cmb_TipDoc_Emp.ItemData(cmb_TipDoc_Emp.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & txt_NumDoc_Emp.Text & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      pnl_FlgEmp.Caption = moddat_gf_Consulta_ParDes("019", "9")
      pnl_FlgEmp.Tag = "9"
      
      txt_RazSoc_Emp.Enabled = True
      txt_NomCom_Emp.Enabled = True
      txt_Telef1_Emp.Enabled = True
      txt_Telef2_Emp.Enabled = True
      
      txt_RazSoc_Emp.Text = ""
      txt_NomCom_Emp.Text = ""
      txt_Telef1_Emp.Text = ""
      txt_Telef2_Emp.Text = ""
      
      Call gs_SetFocus(txt_RazSoc_Emp)
   Else
      g_rst_Princi.MoveFirst
   
      pnl_FlgEmp.Caption = moddat_gf_Consulta_ParDes("019", g_rst_Princi!DATGEN_CLASIF)
      pnl_FlgEmp.Tag = CStr(g_rst_Princi!DATGEN_CLASIF)
      
      txt_RazSoc_Emp.Text = Trim(g_rst_Princi!DATGEN_RAZSOC)
      txt_NomCom_Emp.Text = Trim(g_rst_Princi!DATGEN_NOMCOM)
      
      txt_Telef1_Emp.Text = Trim(g_rst_Princi!DATGEN_TELEF1 & "")
      txt_Telef2_Emp.Text = Trim(g_rst_Princi!DATGEN_TELEF2 & "")
         
      txt_RazSoc_Emp.Enabled = False
      txt_NomCom_Emp.Enabled = False
      txt_Telef1_Emp.Enabled = False
      txt_Telef2_Emp.Enabled = False
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_DirCas_Click()
   cmb_TipVia.ListIndex = frm_MntCli_02.cmb_TipVia.ListIndex
   txt_NomVia.Text = frm_MntCli_02.txt_NomVia.Text
   txt_NumVia.Text = frm_MntCli_02.txt_Numero.Text
   txt_IntDpt.Text = frm_MntCli_02.txt_Interi.Text
   cmb_TipZon.ListIndex = frm_MntCli_02.cmb_TipZon.ListIndex
   txt_NomZon.Text = frm_MntCli_02.txt_NomZon.Text
   
   cmb_DptDir.ListIndex = frm_MntCli_02.cmb_DptDir.ListIndex
   
   Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
   cmb_PrvDir.ListIndex = frm_MntCli_02.cmb_PrvDir.ListIndex
         
   Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
   cmb_DstDir.ListIndex = frm_MntCli_02.cmb_DstDir.ListIndex
   
   txt_Refere.Text = frm_MntCli_02.txt_Refere.Text
   
   txt_Telef1.Text = frm_MntCli_02.txt_Telefo.Text
   
   Call gs_SetFocus(cmb_CodCiu)
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el N�mero de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 7 Then
      If Len(Trim(txt_NumDoc.Text)) <> 11 Then
         MsgBox "El N�mero de Documento ingresado no es correcto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
   End If

   If cmb_TipVia.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de V�a.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipVia)
      Exit Sub
   End If

   If Len(Trim(txt_NomVia.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre de V�a.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomVia)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumVia.Text)) = 0 Then
      MsgBox "Debe ingresar el N�mero.", vbExclamation, modgen_g_str_NomPlt
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
      MsgBox "Debe seleccionar el Departamento de la Direcci�n.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DptDir)
      Exit Sub
   End If
   
   If cmb_PrvDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Provincia de la Direcci�n.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PrvDir)
      Exit Sub
   End If
   
   If cmb_DstDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Distrito de la Direcci�n.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DstDir)
      Exit Sub
   End If

   If Len(Trim(txt_Telef1.Text)) = 0 Then
      MsgBox "Debe ingresar el Tel�fono.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Telef1)
      Exit Sub
   End If

   If cmb_CodCiu.ListIndex = -1 Then
      MsgBox "Debe seleccionar el CIIU.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodCiu)
      Exit Sub
   End If

   If ipp_IngNet.Value = 0 Then
      MsgBox "El Ingreso Declarado no puede ser igual a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_IngNet)
      Exit Sub
   End If
   
   If CDate(ipp_IniAct.Text) > Date Then
      MsgBox "La Fecha de Inicio de Actividades no puede ser mayor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_IniAct)
      Exit Sub
   End If
   
   If cmb_ConLoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar si tiene Contrato de Locaci�n de Servicios.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ConLoc)
      Exit Sub
   End If
   
   If cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex) = 1 Then
      If cmb_TipDoc_Emp.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipDoc_Emp)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumDoc_Emp.Text)) = 0 Then
         MsgBox "Debe ingresar el N�mero de Documento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc_Emp)
         Exit Sub
      End If
      
      If cmb_TipDoc_Emp.ItemData(cmb_TipDoc_Emp.ListIndex) = 7 Then
         If Len(Trim(txt_NumDoc_Emp.Text)) <> 11 Then
            MsgBox "El N�mero de Documento ingresado no es correcto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumDoc_Emp)
            Exit Sub
         End If
      End If
   
      If Len(Trim(txt_RazSoc_Emp.Text)) = 0 Then
         MsgBox "Debe ingresar la Raz�n Social.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_RazSoc_Emp)
         Exit Sub
      End If
   
      If Len(Trim(txt_NomCom_Emp.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre Comercial.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomCom_Emp)
         Exit Sub
      End If
   
      If Len(Trim(txt_Telef1_Emp.Text)) = 0 Then
         MsgBox "Debe ingresar el Tel�fono.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Telef1_Emp)
         Exit Sub
      End If
   
      If CDate(ipp_FecIng_Emp.Text) > CDate(Date) Then
         MsgBox "La Fecha de Ingreso no puede ser mayor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIng_Emp)
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
   End If
   

   If MsgBox("�Est� seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_Inicia_ActEco(moddat_g_int_TipCli, moddat_g_int_OrdAct)
   
   If moddat_g_int_TipCli = 1 Then
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_OrdAct = moddat_g_int_OrdAct
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_TipAct = 21
      
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NumDoc = txt_NumDoc.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_TipVia = cmb_TipVia.ItemData(cmb_TipVia.ListIndex)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NomVia = txt_NomVia.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NumVia = txt_NumVia.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_IntDpt = txt_IntDpt.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_TipZon = cmb_TipZon.ItemData(cmb_TipZon.ListIndex)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NomZon = txt_NomZon.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo = Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00")
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_Refere = txt_Refere.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_Telef1 = txt_Telef1.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_Telef2 = txt_Telef2.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NumFax = txt_NumFax.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_CodCiu = cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)
      
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_IngNet = CDbl(ipp_IngNet.Text)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_IniAct = ipp_IniAct.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_ConLoc = cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex)
      
      If cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex) = 1 Then
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_TipDoc_Emp = cmb_TipDoc_Emp.ItemData(cmb_TipDoc_Emp.ListIndex)
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NumDoc_Emp = txt_NumDoc_Emp.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_FlgEmp = pnl_FlgEmp.Tag
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_RazSoc_Emp = txt_RazSoc_Emp.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NomCom_Emp = txt_NomCom_Emp.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_Telef1_Emp = txt_Telef1_Emp.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_Telef2_Emp = txt_Telef2_Emp.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_FecIng_Emp = ipp_FecIng_Emp.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_CodCar = l_arr_NomCar(cmb_NomCar.ListIndex + 1).Genera_Codigo
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NomCar = txt_NomCar.Text
      End If
   Else
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_OrdAct = moddat_g_int_OrdAct
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_TipAct = 21
      
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NumDoc = txt_NumDoc.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_TipVia = cmb_TipVia.ItemData(cmb_TipVia.ListIndex)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NomVia = txt_NomVia.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NumVia = txt_NumVia.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_IntDpt = txt_IntDpt.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_TipZon = cmb_TipZon.ItemData(cmb_TipZon.ListIndex)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NomZon = txt_NomZon.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo = Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00")
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_Refere = txt_Refere.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_Telef1 = txt_Telef1.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_Telef2 = txt_Telef2.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NumFax = txt_NumFax.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_CodCiu = cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)
      
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_IngNet = CDbl(ipp_IngNet.Text)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_IniAct = ipp_IniAct.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_ConLoc = cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex)
      
      If cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex) = 1 Then
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_TipDoc_Emp = cmb_TipDoc_Emp.ItemData(cmb_TipDoc_Emp.ListIndex)
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NumDoc_Emp = txt_NumDoc_Emp.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_FlgEmp = pnl_FlgEmp.Tag
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_RazSoc_Emp = txt_RazSoc_Emp.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NomCom_Emp = txt_NomCom_Emp.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_Telef1_Emp = txt_Telef1_Emp.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_Telef2_Emp = txt_Telef2_Emp.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_FecIng_Emp = ipp_FecIng_Emp.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_CodCar = l_arr_NomCar(cmb_NomCar.ListIndex + 1).Genera_Codigo
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NomCar = txt_NomCar.Text
      End If
   End If
   
   moddat_g_int_FlgAct_1 = 2
   Unload Me
End Sub

Private Sub cmd_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opci�n.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_11.Show 1
End Sub

Private Sub ipp_FecIng_Emp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_NomCar)
   End If
End Sub

Private Sub txt_NomCar_GotFocus()
   Call gs_SelecTodo(txt_NomCar)
End Sub

Private Sub txt_NomCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
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

Private Sub cmb_DstDir_LostFocus()
   Call SendMessage(cmb_DstDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
      
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   If moddat_g_int_TipCli = 1 Then
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
      
      If moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_TipAct = 21 Then
         Call gs_BuscarCombo_Item(cmb_TipDoc, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_TipDoc)
         txt_NumDoc.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NumDoc
         
         Call gs_BuscarCombo_Item(cmb_TipVia, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_TipVia)
         txt_NomVia.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NomVia
         txt_NumVia.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NumVia
         txt_IntDpt.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_IntDpt
         Call gs_BuscarCombo_Item(cmb_TipZon, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_TipZon)
         txt_NomZon.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NomZon
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir, Left(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 2), Mid(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 2)))
         txt_Refere.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_Refere
         txt_Telef1.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_Telef1
         txt_Telef2.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_Telef2
         txt_NumFax.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NumFax
         
         Call gs_BuscarCombo_Item(cmb_CodCiu, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_CodCiu)
         
         ipp_IngNet.Value = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_IngNet
         ipp_IniAct.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_IniAct
         
         Call gs_BuscarCombo_Item(cmb_ConLoc, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_ConLoc)
         
         If cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex) = 1 Then
            Call gs_BuscarCombo_Item(cmb_TipDoc_Emp, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_TipDoc_Emp)
            txt_NumDoc_Emp.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NumDoc_Emp
            
            pnl_FlgEmp.Tag = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_FlgEmp
            pnl_FlgEmp.Caption = moddat_gf_Consulta_ParDes("019", pnl_FlgEmp.Tag)
            
            txt_RazSoc_Emp.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_RazSoc_Emp
            txt_NomCom_Emp.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NomCom_Emp
            
            txt_Telef1_Emp.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_Telef1_Emp
            txt_Telef2_Emp.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_Telef2_Emp
         
            ipp_FecIng_Emp.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_FecIng_Emp
            
            cmb_NomCar.ListIndex = gf_Busca_Arregl(l_arr_NomCar, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_CodCar) - 1
            txt_NomCar.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ind_NomCar
         End If
      End If
   Else
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli & "(" & CStr(moddat_g_int_CygTDo) & " - " & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom & ")"
   
      If moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_TipAct = 21 Then
         Call gs_BuscarCombo_Item(cmb_TipDoc, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_TipDoc)
         txt_NumDoc.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NumDoc
         
         Call gs_BuscarCombo_Item(cmb_TipVia, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_TipVia)
         txt_NomVia.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NomVia
         txt_NumVia.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NumVia
         txt_IntDpt.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_IntDpt
         Call gs_BuscarCombo_Item(cmb_TipZon, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_TipZon)
         txt_NomZon.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NomZon
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir, Left(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 2), Mid(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_UbiGeo, 2)))
         txt_Refere.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_Refere
         txt_Telef1.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_Telef1
         txt_Telef2.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_Telef2
         txt_NumFax.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NumFax
         
         Call gs_BuscarCombo_Item(cmb_CodCiu, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_CodCiu)
         
         ipp_IngNet.Value = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_IngNet
         ipp_IniAct.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_IniAct
         
         Call gs_BuscarCombo_Item(cmb_ConLoc, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_ConLoc)
         
         If cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex) = 1 Then
            Call gs_BuscarCombo_Item(cmb_TipDoc_Emp, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_TipDoc_Emp)
            txt_NumDoc_Emp.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NumDoc_Emp
            
            pnl_FlgEmp.Tag = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_FlgEmp
            pnl_FlgEmp.Caption = moddat_gf_Consulta_ParDes("019", pnl_FlgEmp.Tag)
            
            txt_RazSoc_Emp.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_RazSoc_Emp
            txt_NomCom_Emp.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NomCom_Emp
            
            txt_Telef1_Emp.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_Telef1_Emp
            txt_Telef2_Emp.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_Telef2_Emp
         
            ipp_FecIng_Emp.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_FecIng_Emp
            
            cmb_NomCar.ListIndex = gf_Busca_Arregl(l_arr_NomCar, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_CodCar) - 1
            txt_NomCar.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ind_NomCar
         End If
      End If
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "232")

   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   Call moddat_gs_Carga_CdCIIU(cmb_CodCiu)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_ConLoc, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Emp, 1, "232")
   
   Call moddat_gs_Carga_LisIte(cmb_NomCar, l_arr_NomCar, 1, "503")
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   
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
   cmb_CodCiu.ListIndex = -1
   
   ipp_IngNet.Value = 0
   ipp_IniAct.Text = Format(Date, "dd/mm/yyyy")
   
   cmb_ConLoc.ListIndex = -1
   
   cmb_TipDoc_Emp.ListIndex = -1
   txt_NumDoc_Emp.Text = ""
   pnl_FlgEmp.Caption = ""
   pnl_FlgEmp.Tag = ""
   txt_RazSoc_Emp.Text = ""
   txt_NomCom_Emp.Text = ""
   txt_Telef1_Emp.Text = ""
   txt_Telef2_Emp.Text = ""
   ipp_FecIng_Emp.Text = Format(Date, "dd/mm/yyyy")
   cmb_NomCar.ListIndex = -1
   txt_NomCar.Text = ""
   txt_NomCar.Enabled = False
   
   cmb_TipDoc_Emp.Enabled = False
   txt_NumDoc_Emp.Enabled = False
   cmd_BusEmp_Emp.Enabled = False
   txt_RazSoc_Emp.Enabled = False
   txt_NomCom_Emp.Enabled = False
   txt_Telef1_Emp.Enabled = False
   txt_Telef2_Emp.Enabled = False
   ipp_FecIng_Emp.Enabled = False
   cmb_NomCar.Enabled = False
End Sub

Private Sub ipp_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IniAct)
   End If
End Sub

Private Sub ipp_IniAct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ConLoc)
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_NomZon_GotFocus()
   Call gs_SelecTodo(txt_NomZon)
End Sub

Private Sub txt_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_NumDoc_Emp_Change()
   pnl_FlgEmp.Caption = ""
   pnl_FlgEmp.Tag = ""
   
   txt_RazSoc_Emp.Text = ""
   txt_NomCom_Emp.Text = ""
   txt_Telef1_Emp.Text = ""
   txt_Telef2_Emp.Text = ""
   
   txt_RazSoc_Emp.Enabled = False
   txt_NomCom_Emp.Enabled = False
   txt_Telef1_Emp.Enabled = False
   txt_Telef2_Emp.Enabled = False
End Sub

Private Sub txt_NumDoc_Emp_GotFocus()
   Call gs_SelecTodo(txt_NumDoc_Emp)
End Sub

Private Sub txt_NumDoc_Emp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(pnl_FlgEmp.Caption)) > 0 Then
         Call gs_SetFocus(ipp_FecIng_Emp)
      Else
         Call gs_SetFocus(cmd_BusEmp_Emp)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumVia_GotFocus()
   Call gs_SelecTodo(txt_NumVia)
End Sub

Private Sub txt_NumVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_IntDpt_GotFocus()
   Call gs_SelecTodo(txt_IntDpt)
End Sub

Private Sub txt_IntDpt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_RazSoc_Emp_GotFocus()
   Call gs_SelecTodo(txt_RazSoc_Emp)
End Sub

Private Sub txt_RazSoc_Emp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomCom_Emp)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_NomCom_Emp_GotFocus()
   Call gs_SelecTodo(txt_NomCom_Emp)
End Sub

Private Sub txt_NomCom_Emp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1_Emp)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_Telef1_Emp_GotFocus()
   Call gs_SelecTodo(txt_Telef1_Emp)
End Sub

Private Sub txt_Telef1_Emp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef2_Emp)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Telef2_Emp_GotFocus()
   Call gs_SelecTodo(txt_Telef2_Emp)
End Sub

Private Sub txt_Telef2_Emp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIng_Emp)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?��")
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
      Call gs_SetFocus(cmb_CodCiu)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

