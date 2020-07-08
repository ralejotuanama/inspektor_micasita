VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_TraCof_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9615
   ClientLeft      =   225
   ClientTop       =   465
   ClientWidth     =   13425
   Icon            =   "AteCli_frm_032.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9615
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   13425
      _Version        =   65536
      _ExtentX        =   23680
      _ExtentY        =   16960
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
         Height          =   7125
         Left            =   30
         TabIndex        =   27
         Top             =   1590
         Width           =   13335
         _Version        =   65536
         _ExtentX        =   23521
         _ExtentY        =   12568
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
         Begin VB.ComboBox cmb_CtaBan 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2040
            Width           =   3315
         End
         Begin VB.CommandButton cmd_BusSol 
            Caption         =   "..."
            Height          =   315
            Left            =   3300
            TabIndex        =   20
            Top             =   5550
            Width           =   465
         End
         Begin VB.ComboBox cmb_CodBan 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1710
            Width           =   3315
         End
         Begin VB.CommandButton cmd_Cancel 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   1830
            TabIndex        =   22
            Top             =   6690
            Width           =   1755
         End
         Begin VB.CommandButton cmd_Agrega 
            Caption         =   "&Agregar a Lista"
            Height          =   375
            Left            =   60
            TabIndex        =   21
            Top             =   6690
            Width           =   1755
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1620
            MaxLength       =   120
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   5550
            Width           =   1635
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   5220
            Width           =   3315
         End
         Begin VB.TextBox txt_NumOpe 
            Height          =   315
            Left            =   1620
            MaxLength       =   120
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   4890
            Width           =   3315
         End
         Begin VB.CommandButton cmd_NueOpe 
            Caption         =   "&Nueva Operación"
            Height          =   375
            Left            =   11490
            TabIndex        =   13
            Top             =   3270
            Width           =   1755
         End
         Begin VB.CommandButton cmd_EdiOpe 
            Caption         =   "&Editar Operación"
            Height          =   375
            Left            =   11490
            TabIndex        =   14
            Top             =   3660
            Width           =   1755
         End
         Begin VB.CommandButton cmd_BorOpe 
            Caption         =   "&Borrar Operación"
            Height          =   375
            Left            =   11490
            TabIndex        =   15
            Top             =   4050
            Width           =   1755
         End
         Begin VB.TextBox txt_Observ 
            Height          =   465
            Left            =   1620
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Text            =   "AteCli_frm_032.frx":000C
            Top             =   2370
            Width           =   11625
         End
         Begin VB.TextBox txt_TipOpe 
            Height          =   315
            Left            =   1620
            MaxLength       =   120
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   720
            Width           =   3315
         End
         Begin EditLib.fpDateTime ipp_FecEmi 
            Height          =   315
            Left            =   1620
            TabIndex        =   4
            Top             =   60
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
         Begin EditLib.fpDateTime ipp_FecVal 
            Height          =   315
            Left            =   1620
            TabIndex        =   5
            Top             =   390
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
         Begin EditLib.fpDoubleSingle ipp_ImpTot 
            Height          =   315
            Left            =   1620
            TabIndex        =   7
            Top             =   1050
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
         Begin Threed.SSPanel SSPanel22 
            Height          =   90
            Left            =   30
            TabIndex        =   35
            Top             =   2880
            Width           =   13275
            _Version        =   65536
            _ExtentX        =   23416
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1455
            Left            =   30
            TabIndex        =   12
            Top             =   3270
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   21
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   1530
            TabIndex        =   36
            Top             =   3000
            Width           =   6075
            _Version        =   65536
            _ExtentX        =   10716
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cliente"
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   60
            TabIndex        =   37
            Top             =   3000
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operación"
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
         Begin Threed.SSPanel SSPanel20 
            Height          =   285
            Left            =   9390
            TabIndex        =   38
            Top             =   3000
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   7590
            TabIndex        =   39
            Top             =   3000
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
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
            Height          =   90
            Left            =   30
            TabIndex        =   40
            Top             =   4770
            Width           =   13275
            _Version        =   65536
            _ExtentX        =   23416
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1620
            TabIndex        =   44
            Top             =   5880
            Width           =   11625
            _Version        =   65536
            _ExtentX        =   20505
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
         Begin EditLib.fpDoubleSingle ipp_Import 
            Height          =   315
            Left            =   9090
            TabIndex        =   17
            Top             =   4890
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   47
            Top             =   6210
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "001-001-04-0001"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   90
            Left            =   30
            TabIndex        =   49
            Top             =   6570
            Width           =   13275
            _Version        =   65536
            _ExtentX        =   23416
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
         Begin VB.Label Label13 
            Caption         =   "Nombre Banco:"
            Height          =   315
            Left            =   60
            TabIndex        =   50
            Top             =   1710
            Width           =   1275
         End
         Begin VB.Label Label12 
            Caption         =   "Número Solicitud:"
            Height          =   315
            Left            =   60
            TabIndex        =   48
            Top             =   6210
            Width           =   1275
         End
         Begin VB.Label Label11 
            Caption         =   "Importe:"
            Height          =   285
            Left            =   7560
            TabIndex        =   46
            Top             =   4890
            Width           =   1485
         End
         Begin VB.Label Label10 
            Caption         =   "Nombre Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   45
            Top             =   5880
            Width           =   1275
         End
         Begin VB.Label Label9 
            Caption         =   "Número Doc. Ident.:"
            Height          =   285
            Left            =   60
            TabIndex        =   43
            Top             =   5550
            Width           =   1485
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   60
            TabIndex        =   42
            Top             =   5220
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "Número Operación:"
            Height          =   285
            Left            =   60
            TabIndex        =   41
            Top             =   4890
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Observaciones:"
            Height          =   285
            Left            =   60
            TabIndex        =   34
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Nro. Cuenta Abono:"
            Height          =   285
            Left            =   60
            TabIndex        =   33
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Operación:"
            Height          =   285
            Left            =   60
            TabIndex        =   32
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Importe Total:"
            Height          =   285
            Left            =   90
            TabIndex        =   31
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Moneda:"
            Height          =   315
            Left            =   60
            TabIndex        =   30
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Valor:"
            Height          =   315
            Left            =   60
            TabIndex        =   29
            Top             =   390
            Width           =   1365
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Emisión:"
            Height          =   315
            Left            =   60
            TabIndex        =   28
            Top             =   60
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   13335
         _Version        =   65536
         _ExtentX        =   23521
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
            TabIndex        =   26
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Trámites COFIDE - Recepción de Carta"
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
            Picture         =   "AteCli_frm_032.frx":0010
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   795
         Left            =   30
         TabIndex        =   51
         Top             =   750
         Width           =   13335
         _Version        =   65536
         _ExtentX        =   23521
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   11130
            Picture         =   "AteCli_frm_032.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   11850
            Picture         =   "AteCli_frm_032.frx":0624
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12570
            Picture         =   "AteCli_frm_032.frx":092E
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   60
            Width           =   675
         End
         Begin VB.TextBox txt_NumCar 
            Height          =   315
            Left            =   1620
            MaxLength       =   12
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   53
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label Label18 
            Caption         =   "Nro. Carta:"
            Height          =   315
            Left            =   60
            TabIndex        =   52
            Top             =   240
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   30
         TabIndex        =   54
         Top             =   8760
         Width           =   13335
         _Version        =   65536
         _ExtentX        =   23521
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
            Left            =   12630
            Picture         =   "AteCli_frm_032.frx":0D70
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_TraCof_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CodBan()   As moddat_tpo_Genera
Dim l_arr_CtaBan()   As moddat_tpo_Genera
Dim l_str_EmiCar     As String
Dim l_str_RegCar     As String

Private Sub cmb_CodBan_Click()
   If cmb_CodBan.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_CtaBan(l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, cmb_CtaBan, l_arr_CtaBan)
      Screen.MousePointer = 0
         
      Call gs_SetFocus(cmb_CtaBan)
   Else
      cmb_CtaBan.Clear
   End If
End Sub

Private Sub cmb_CodBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodBan_Click
   End If
End Sub

Private Sub cmb_CtaBan_Click()
   Call gs_SetFocus(txt_Observ)
End Sub

Private Sub cmb_CtaBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaBan_Click
   End If
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(ipp_ImpTot)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   Dim r_int_Contad  As Integer
   
   If Len(Trim(txt_NumOpe.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumOpe)
      Exit Sub
   End If
   
   If ipp_Import.Value = 0 Then
      MsgBox "Debe ingresar el Importe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import)
      Exit Sub
   End If
   
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

   If Len(Trim(pnl_NumSol.Caption)) = 0 Then
      MsgBox "Debe ubicar la Solicitud para el Cliente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_BusSol)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         grd_Listad.Col = 0
         
         If grd_Listad.Text = txt_NumOpe.Text Then
            Call gs_UbiIniGrid(grd_Listad)
            
            MsgBox "Ya ingreso este Número de Operación.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumOpe)
            
            Exit Sub
         End If
      
         grd_Listad.Col = 1
         If pnl_NomCli.Caption = grd_Listad.Text Then
            Call gs_UbiIniGrid(grd_Listad)
            
            MsgBox "Ya registro a este Cliente.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumDoc)
            
            Exit Sub
         End If
      Next r_int_Contad
   End If
   
   If MsgBox("¿Está seguro de agregar la Operación a la Lista?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
   End If
   
   grd_Listad.Col = 0
   grd_Listad.Text = txt_NumOpe.Text
   
   grd_Listad.Col = 1
   grd_Listad.Text = pnl_NomCli.Caption
   
   grd_Listad.Col = 2
   grd_Listad.Text = pnl_NumSol.Caption
   
   grd_Listad.Col = 3
   grd_Listad.Text = ipp_Import.Text
   
   grd_Listad.Col = 4
   grd_Listad.Text = txt_NumDoc.Text
   
   grd_Listad.Col = 5
   grd_Listad.Text = moddat_g_str_NomCli
   
   Call cmd_Cancel_Click
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Sub cmd_BorOpe_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   If MsgBox("¿Está seguro de eliminar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If grd_Listad.Rows > 1 Then
      grd_Listad.RemoveItem (grd_Listad.Row)
   Else
      Call gs_LimpiaGrid(grd_Listad)
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If Len(Trim(txt_NumCar.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Carta COFIDE.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumCar)
      Exit Sub
   End If
   
   g_str_Parame = "SELECT * FROM TRA_CARCOF WHERE "
   g_str_Parame = g_str_Parame & "CARCOF_NUMCAR = '" & txt_NumCar.Text & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
      
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      MsgBox "La Carta ya fue registrada.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumCar)
      Exit Sub
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call fs_Activa(True)
   Call gs_SetFocus(ipp_FecEmi)
End Sub

Private Sub cmd_BusSol_Click()
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
   
   moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   moddat_g_str_NumDoc = txt_NumDoc.Text

   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = '" & moddat_g_str_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_ENVCRE = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "El Cliente no presenta ninguna Solicitud en Trámite. ", vbExclamation, modgen_g_str_NomPlt
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If

   moddat_g_str_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   moddat_g_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   If moddat_g_str_CodPrd <> "001" Then
      MsgBox "Solo se realizan Trámites co COFIDE para el Producto Mivivienda.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If

   'Validación que se encuentre en Instancia
   If Not (moddat_g_int_InsAct = modatecli_g_con_PolSeg Or moddat_g_int_InsAct = modatecli_g_con_TraCof) Then
      MsgBox "No se encuentra en Instancia de Trámites COFIDE.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   l_str_RegCar = ""
   l_str_EmiCar = ""

   'Obteniendo Información del Seguimiento
   Call fs_Buscar_SegDet
   
   If Len(Trim(l_str_EmiCar)) = 0 Then
      MsgBox "No se ha emitido Cartas de Solicitud de Fondos.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If

   If Len(Trim(l_str_RegCar)) > 0 Then
      MsgBox "Ya se ha registrado Carta COFIDE para este Cliente.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If

   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)

   Call gs_SetFocus(cmd_Agrega)
End Sub

Private Sub fs_Buscar_SegDet()
   Dim r_str_FecOcu  As String
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(modatecli_g_con_TraCof) & " "
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
         Case 71:    l_str_EmiCar = r_str_FecOcu
         Case 72:    l_str_RegCar = r_str_FecOcu
      End Select
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_LimpiaItem
   Call fs_ActivaItem(False)
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_EdiOpe_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   txt_NumOpe.Text = grd_Listad.Text
   
   grd_Listad.Col = 3
   ipp_Import.Value = CDbl(grd_Listad.Text)
   
   grd_Listad.Col = 1
   Call gs_BuscarCombo_Item(cmb_TipDoc, CInt(Left(grd_Listad.Text, 1)))
   
   pnl_NomCli.Caption = grd_Listad.Text
   
   grd_Listad.Col = 4
   txt_NumDoc.Text = grd_Listad.Text
      
   grd_Listad.Col = 2
   pnl_NumSol.Caption = grd_Listad.Text
      
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   Call fs_ActivaItem(True)
   
   txt_NumOpe.Enabled = False
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   cmd_BusSol.Enabled = False
   
   Call gs_SetFocus(ipp_Import)
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_Contad     As Integer
   Dim r_dbl_TotCar     As Double
   Dim r_str_NumOpe     As String
   Dim r_dbl_Import     As Double
   Dim r_str_NumSol     As String
   
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
   If ipp_ImpTot.Value = 0 Then
      MsgBox "Debe ingresar el Importe Total de la Carta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ImpTot)
      Exit Sub
   End If
   
   If Len(Trim(txt_TipOpe.Text)) = 0 Then
      MsgBox "Debe ingresar el Tipo de Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_TipOpe)
      Exit Sub
   End If
   
   If cmb_CodBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodBan)
      Exit Sub
   End If
   
   If cmb_CtaBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Número de Cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaBan)
      Exit Sub
   End If
   
   If grd_Listad.Rows = 0 Then
      MsgBox "Debe ingresar las Operaciones.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_NueOpe)
      Exit Sub
   End If
   
   'Validar que el Total Cuadre con la suma
   r_dbl_TotCar = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      grd_Listad.Col = 3
      
      r_dbl_TotCar = r_dbl_TotCar + CDbl(grd_Listad.Text)
   Next r_int_Contad
   
   Call gs_UbiIniGrid(grd_Listad)
   
   If r_dbl_TotCar <> ipp_ImpTot.Value Then
      MsgBox "No cuadra la Suma de las Operaciones con el Total de la Carta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ImpTot)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando Cabecera de Carta
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_CARCOF ("
   
      g_str_Parame = g_str_Parame & "'" & txt_NumCar.Text & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecEmi.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecVal.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ImpTot.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_TipOpe.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
      
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_TRA_CARCOF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 0
      r_str_NumOpe = grd_Listad.Text
      
      grd_Listad.Col = 2
      r_str_NumSol = Left(grd_Listad.Text, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
   
      grd_Listad.Col = 3
      r_dbl_Import = CDbl(grd_Listad.Text)
   
   
      'Grabando Detalle de Carta
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_TRA_DETCOF_CARCOF ("
      
         g_str_Parame = g_str_Parame & "'" & r_str_NumSol & "', "
         g_str_Parame = g_str_Parame & "'" & txt_NumCar.Text & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_NumOpe & "', "
         g_str_Parame = g_str_Parame & CStr(r_dbl_Import) & ", "
         g_str_Parame = g_str_Parame & "'" & l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo & "', "
            
         'Datos de Auditoria
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
         g_str_Parame = g_str_Parame & "1)"
            
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
   
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento USP_TRA_DETCOF_CARCOF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   
      If Not moddat_gf_Inserta_SegDet(r_str_NumSol, modatecli_g_con_TraCof, 72, 0, "", 0, 0) Then
         Exit Sub
      End If
   Next r_int_Contad

   MsgBox "Se registraron los datos correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Call cmd_Limpia_Click
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(txt_NumCar)
End Sub

Private Sub cmd_NueOpe_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_ActivaItem(True)
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

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_FecEmi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecVal)
   End If
End Sub

Private Sub ipp_FecVal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipMon)
   End If
End Sub

Private Sub ipp_Import_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc)
   End If
End Sub

Private Sub ipp_ImpTot_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_TipOpe)
   End If
End Sub

Private Sub txt_NumCar_GotFocus()
   Call gs_SelecTodo(txt_NumCar)
End Sub

Private Sub txt_NumCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_.:,;")
   End If
End Sub

Private Sub txt_NumOpe_GotFocus()
   Call gs_SelecTodo(txt_NumOpe)
End Sub

Private Sub txt_NumOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Import)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_.:()")
   End If
End Sub

Private Sub txt_TipOpe_GotFocus()
   Call gs_SelecTodo(txt_TipOpe)
End Sub

Private Sub txt_TipOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodBan)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_.,:;()=?¿%&/")
   End If
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_NueOpe)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1475
   grd_Listad.ColWidth(1) = 6065
   grd_Listad.ColWidth(2) = 1805
   grd_Listad.ColWidth(3) = 1660
   grd_Listad.ColWidth(4) = 0
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColWidth(6) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter

   Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)
   Call moddat_gs_Carga_TipMon(cmb_TipMon, 1)
   
   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "505")
End Sub

Private Sub fs_Limpia()
   txt_NumCar.Text = ""
   ipp_FecEmi.Text = Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy")
   ipp_FecVal.Text = Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy")
   
   cmb_TipMon.ListIndex = -1
   ipp_ImpTot.Value = 0
   txt_TipOpe.Text = ""
   cmb_CodBan.ListIndex = -1
   cmb_CtaBan.Clear
   txt_Observ.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_LimpiaItem
   
   Call fs_ActivaItem(False)
   Call fs_Activa(False)
End Sub

Private Sub fs_LimpiaItem()
   txt_NumOpe.Text = ""
   ipp_Import.Value = 0
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   pnl_NomCli.Caption = ""
   pnl_NumSol.Caption = ""
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

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusSol)
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

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   txt_NumCar.Enabled = Not p_Habilita
   cmd_Buscar.Enabled = Not p_Habilita

   ipp_FecEmi.Enabled = p_Habilita
   ipp_FecVal.Enabled = p_Habilita
   cmb_TipMon.Enabled = p_Habilita
   ipp_ImpTot.Enabled = p_Habilita
   txt_TipOpe.Enabled = p_Habilita
   cmb_CodBan.Enabled = p_Habilita
   cmb_CtaBan.Enabled = p_Habilita
   txt_Observ.Enabled = p_Habilita
   
   grd_Listad.Enabled = p_Habilita
   cmd_NueOpe.Enabled = p_Habilita
   cmd_EdiOpe.Enabled = p_Habilita
   cmd_BorOpe.Enabled = p_Habilita
   cmd_Grabar.Enabled = p_Habilita
End Sub

Private Sub fs_ActivaItem(ByVal p_Habilita As Integer)
   txt_NumOpe.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   ipp_Import.Enabled = p_Habilita
   cmd_BusSol.Enabled = p_Habilita
   cmd_Agrega.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   
   ipp_FecEmi.Enabled = Not p_Habilita
   ipp_FecVal.Enabled = Not p_Habilita
   cmb_TipMon.Enabled = Not p_Habilita
   ipp_ImpTot.Enabled = Not p_Habilita
   txt_TipOpe.Enabled = Not p_Habilita
   cmb_CodBan.Enabled = Not p_Habilita
   cmb_CtaBan.Enabled = Not p_Habilita
   txt_Observ.Enabled = Not p_Habilita
   
   grd_Listad.Enabled = Not p_Habilita
   cmd_NueOpe.Enabled = Not p_Habilita
   cmd_EdiOpe.Enabled = Not p_Habilita
   cmd_BorOpe.Enabled = Not p_Habilita
   cmd_Grabar.Enabled = Not p_Habilita
End Sub
