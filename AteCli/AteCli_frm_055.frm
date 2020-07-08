VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_IngSol_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10020
   ClientLeft      =   2955
   ClientTop       =   465
   ClientWidth     =   11610
   Icon            =   "AteCli_frm_055.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10005
      Left            =   0
      TabIndex        =   144
      Top             =   0
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
      _ExtentY        =   17648
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
         Height          =   465
         Left            =   30
         TabIndex        =   195
         Top             =   8640
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   820
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
         Begin VB.CommandButton cmd_Agrega 
            Caption         =   "&Agregar a Lista"
            Height          =   345
            Left            =   60
            TabIndex        =   153
            Top             =   60
            Width           =   1755
         End
         Begin VB.CommandButton cmd_Cancel 
            Caption         =   "&Cancelar"
            Height          =   345
            Left            =   1860
            TabIndex        =   155
            Top             =   60
            Width           =   1755
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   795
         Left            =   30
         TabIndex        =   194
         Top             =   9150
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10800
            Picture         =   "AteCli_frm_055.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   159
            ToolTipText     =   "Salir de la Opción"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10080
            Picture         =   "AteCli_frm_055.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   157
            ToolTipText     =   "Aceptar Datos"
            Top             =   60
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   158
         Top             =   2220
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.ComboBox cmb_OrdAct 
            Height          =   315
            Left            =   8100
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   60
            Width           =   3315
         End
         Begin VB.ComboBox cmb_ActEco 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label Label19 
            Caption         =   "Actividad Económica:"
            Height          =   315
            Left            =   60
            TabIndex        =   161
            Top             =   60
            Width           =   1905
         End
         Begin VB.Label Label18 
            Caption         =   "Orden Actividad Econom.:"
            Height          =   315
            Left            =   6090
            TabIndex        =   160
            Top             =   60
            Width           =   2115
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   152
         Top             =   750
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.CommandButton cmd_NueAct 
            Caption         =   "&Nueva Actividad"
            Height          =   345
            Left            =   9660
            TabIndex        =   1
            Top             =   330
            Width           =   1755
         End
         Begin VB.CommandButton cmd_BorAct 
            Caption         =   "&Borrar Actividad"
            Height          =   345
            Left            =   9660
            TabIndex        =   2
            Top             =   690
            Width           =   1755
         End
         Begin VB.CommandButton cmd_EdiAct 
            Caption         =   "&Editar Actividad"
            Height          =   345
            Left            =   9660
            TabIndex        =   3
            Top             =   1050
            Width           =   1755
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1035
            Left            =   30
            TabIndex        =   0
            Top             =   330
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   1826
            _Version        =   393216
            Rows            =   12
            Cols            =   100
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   3300
            TabIndex        =   154
            Top             =   60
            Width           =   5955
            _Version        =   65536
            _ExtentX        =   10504
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Actividad Económica"
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
            Left            =   60
            TabIndex        =   156
            Top             =   60
            Width           =   3255
            _Version        =   65536
            _ExtentX        =   5741
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Orden Actividad Económica"
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   30
         TabIndex        =   146
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   495
            Left            =   630
            TabIndex        =   148
            Top             =   60
            Width           =   3195
            _Version        =   65536
            _ExtentX        =   5636
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Actividades Económicas"
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
            Left            =   3720
            TabIndex        =   150
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
            Picture         =   "AteCli_frm_055.frx":0758
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_TraInd 
         Height          =   5895
         Left            =   30
         TabIndex        =   196
         Top             =   2700
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   10398
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
         Begin VB.TextBox txt_Ind_NomCar 
            Height          =   315
            Left            =   8100
            MaxLength       =   250
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   5190
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Ind_NomCar 
            Height          =   315
            Left            =   1980
            TabIndex        =   33
            Text            =   "cmb_Dep_NomCar"
            Top             =   5190
            Width           =   3315
         End
         Begin VB.CommandButton cmd_Ind_BusEmp 
            Caption         =   "..."
            Height          =   315
            Left            =   3780
            TabIndex        =   29
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   4200
            Width           =   435
         End
         Begin VB.TextBox txt_Ind_NDoEmp 
            Height          =   315
            Left            =   1980
            MaxLength       =   11
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   4200
            Width           =   1755
         End
         Begin VB.ComboBox cmb_Ind_ConLoc 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   3540
            Width           =   765
         End
         Begin VB.ComboBox cmb_Ind_TDoEmp 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   3870
            Width           =   2745
         End
         Begin VB.TextBox txt_Ind_RazSoc 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   4530
            Width           =   9435
         End
         Begin VB.TextBox txt_Ind_Tl1Emp 
            Height          =   315
            Left            =   1980
            MaxLength       =   12
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   4860
            Width           =   1365
         End
         Begin VB.TextBox txt_Ind_Tl2Emp 
            Height          =   315
            Left            =   3360
            MaxLength       =   12
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   4860
            Width           =   1365
         End
         Begin VB.CommandButton cmd_Ind_Direcc 
            Caption         =   "="
            Height          =   315
            Left            =   1500
            TabIndex        =   10
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   1080
            Width           =   435
         End
         Begin VB.TextBox txt_Ind_Telef2 
            Height          =   315
            Left            =   3630
            MaxLength       =   12
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   2730
            Width           =   1640
         End
         Begin VB.ComboBox cmb_Ind_TipVia 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1080
            Width           =   3315
         End
         Begin VB.TextBox txt_Ind_Numero 
            Height          =   315
            Left            =   8100
            MaxLength       =   15
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   1410
            Width           =   1640
         End
         Begin VB.TextBox txt_Ind_Interi 
            Height          =   315
            Left            =   9750
            MaxLength       =   15
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   1410
            Width           =   1640
         End
         Begin VB.TextBox txt_Ind_NomZon 
            Height          =   315
            Left            =   8100
            MaxLength       =   120
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   1740
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Ind_PrvDir 
            Height          =   315
            Left            =   8100
            TabIndex        =   18
            Text            =   "cmb_PrvDir"
            Top             =   2070
            Width           =   3315
         End
         Begin VB.TextBox txt_Ind_Refere 
            Height          =   315
            Left            =   8100
            MaxLength       =   250
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   2400
            Width           =   3315
         End
         Begin VB.TextBox txt_Ind_NumFax 
            Height          =   315
            Left            =   8100
            MaxLength       =   12
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   2730
            Width           =   1640
         End
         Begin VB.TextBox txt_Ind_GirCom 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   750
            Width           =   9495
         End
         Begin VB.ComboBox cmb_Ind_TipDoc 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   90
            Width           =   3315
         End
         Begin VB.TextBox txt_Ind_NumDoc 
            Height          =   315
            Left            =   8100
            MaxLength       =   11
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   90
            Width           =   2415
         End
         Begin VB.ComboBox cmb_Ind_GirCom 
            Height          =   315
            Left            =   1980
            TabIndex        =   8
            Text            =   "cmb_GirCom"
            Top             =   420
            Width           =   9495
         End
         Begin VB.TextBox txt_Ind_NomVia 
            Height          =   315
            Left            =   1980
            MaxLength       =   120
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   1410
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Ind_TipZon 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1740
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Ind_DptDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   17
            Text            =   "cmb_DptDir"
            Top             =   2070
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Ind_DstDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   19
            Text            =   "cmb_DstDir"
            Top             =   2400
            Width           =   3315
         End
         Begin VB.TextBox txt_Ind_Telef1 
            Height          =   315
            Left            =   1980
            MaxLength       =   12
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   2730
            Width           =   1640
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   90
            Left            =   30
            TabIndex        =   197
            Top             =   3090
            Width           =   11445
            _Version        =   65536
            _ExtentX        =   20188
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
         Begin EditLib.fpDoubleSingle ipp_Ind_IngNet 
            Height          =   315
            Left            =   1980
            TabIndex        =   24
            Top             =   3210
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
         Begin EditLib.fpDateTime ipp_Ind_FecIng 
            Height          =   315
            Left            =   1980
            TabIndex        =   35
            Top             =   5520
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
            ButtonStyle     =   1
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
         Begin Threed.SSPanel pnl_Ind_FlgEmp 
            Height          =   315
            Left            =   4260
            TabIndex        =   221
            Top             =   4200
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
         Begin EditLib.fpDateTime ipp_Ind_FecIni 
            Height          =   315
            Left            =   8100
            TabIndex        =   25
            Top             =   3210
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
            ButtonStyle     =   1
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
         Begin VB.Label Label5 
            Caption         =   "Fecha Inicio Actividades:"
            Height          =   315
            Left            =   6090
            TabIndex        =   287
            Top             =   3210
            Width           =   1905
         End
         Begin VB.Label lbl_General 
            Caption         =   "Cargo (Especificar):"
            Height          =   285
            Index           =   67
            Left            =   6090
            TabIndex        =   286
            Top             =   5190
            Width           =   2055
         End
         Begin VB.Label Label16 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Left            =   60
            TabIndex        =   220
            Top             =   5520
            Width           =   1605
         End
         Begin VB.Label Label15 
            Caption         =   "Cargo:"
            Height          =   315
            Left            =   60
            TabIndex        =   219
            Top             =   5190
            Width           =   1665
         End
         Begin VB.Label Label14 
            Caption         =   "Razón Social:"
            Height          =   315
            Left            =   60
            TabIndex        =   218
            Top             =   4530
            Width           =   1905
         End
         Begin VB.Label Label13 
            Caption         =   "Nro. Docum. Empresa:"
            Height          =   285
            Left            =   60
            TabIndex        =   217
            Top             =   4200
            Width           =   1755
         End
         Begin VB.Label Label12 
            Caption         =   "Tipo Docum. Empresa:"
            Height          =   315
            Left            =   60
            TabIndex        =   216
            Top             =   3870
            Width           =   1905
         End
         Begin VB.Label Label11 
            Caption         =   "Contrato Locación:"
            Height          =   285
            Left            =   60
            TabIndex        =   215
            Top             =   3540
            Width           =   1785
         End
         Begin VB.Label Label61 
            Caption         =   "Ingresos (S/.):"
            Height          =   285
            Left            =   60
            TabIndex        =   214
            Top             =   3210
            Width           =   2025
         End
         Begin VB.Label Label10 
            Caption         =   "Teléfonos:"
            Height          =   285
            Left            =   60
            TabIndex        =   213
            Top             =   4860
            Width           =   1485
         End
         Begin VB.Label Label63 
            Caption         =   "Tipo de Vía:"
            Height          =   285
            Left            =   60
            TabIndex        =   212
            Top             =   1080
            Width           =   1545
         End
         Begin VB.Label Label62 
            Caption         =   "Giro Comercial (Especif.):"
            Height          =   285
            Left            =   60
            TabIndex        =   211
            Top             =   750
            Width           =   1935
         End
         Begin VB.Label Label60 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   6090
            TabIndex        =   210
            Top             =   1410
            Width           =   1935
         End
         Begin VB.Label Label59 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   6090
            TabIndex        =   209
            Top             =   1740
            Width           =   1485
         End
         Begin VB.Label Label58 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   6090
            TabIndex        =   208
            Top             =   2070
            Width           =   1485
         End
         Begin VB.Label Label57 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   6090
            TabIndex        =   207
            Top             =   2400
            Width           =   1485
         End
         Begin VB.Label Label56 
            Caption         =   "Fax:"
            Height          =   285
            Left            =   6090
            TabIndex        =   206
            Top             =   2730
            Width           =   1485
         End
         Begin VB.Label Label54 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   60
            TabIndex        =   205
            Top             =   1410
            Width           =   1485
         End
         Begin VB.Label Label53 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   60
            TabIndex        =   204
            Top             =   1740
            Width           =   1455
         End
         Begin VB.Label Label52 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   60
            TabIndex        =   203
            Top             =   2070
            Width           =   1425
         End
         Begin VB.Label Label51 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   60
            TabIndex        =   202
            Top             =   2400
            Width           =   1485
         End
         Begin VB.Label Label50 
            Caption         =   "Teléfono:"
            Height          =   285
            Left            =   60
            TabIndex        =   201
            Top             =   2730
            Width           =   1485
         End
         Begin VB.Label Label49 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Left            =   60
            TabIndex        =   200
            Top             =   90
            Width           =   1635
         End
         Begin VB.Label Label48 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Left            =   6090
            TabIndex        =   199
            Top             =   90
            Width           =   1845
         End
         Begin VB.Label Label46 
            Caption         =   "Giro Comercial:"
            Height          =   285
            Left            =   60
            TabIndex        =   198
            Top             =   420
            Width           =   1365
         End
      End
      Begin Threed.SSPanel pnl_TraAcc 
         Height          =   5895
         Left            =   30
         TabIndex        =   250
         Top             =   2700
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   10398
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
         Begin VB.CommandButton cmd_Acc_BusEmp 
            Caption         =   "..."
            Height          =   315
            Left            =   10500
            TabIndex        =   38
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   90
            Width           =   435
         End
         Begin VB.TextBox txt_Acc_Telef1 
            Height          =   315
            Left            =   1980
            MaxLength       =   12
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   3060
            Width           =   1640
         End
         Begin VB.ComboBox cmb_Acc_DstDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   51
            Text            =   "cmb_DstDir"
            Top             =   2730
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Acc_DptDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   49
            Text            =   "cmb_DptDir"
            Top             =   2400
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Acc_TipZon 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   2070
            Width           =   3315
         End
         Begin VB.TextBox txt_Acc_NomVia 
            Height          =   315
            Left            =   1980
            MaxLength       =   120
            TabIndex        =   44
            Text            =   "Text1"
            Top             =   1740
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Acc_GirCom 
            Height          =   315
            Left            =   1980
            TabIndex        =   41
            Text            =   "cmb_GirCom"
            Top             =   750
            Width           =   9465
         End
         Begin VB.TextBox txt_Acc_NumDoc 
            Height          =   315
            Left            =   8100
            MaxLength       =   11
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   90
            Width           =   2355
         End
         Begin VB.ComboBox cmb_Acc_TipDoc 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   90
            Width           =   3315
         End
         Begin VB.TextBox txt_Acc_GirCom 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   1080
            Width           =   9465
         End
         Begin VB.TextBox txt_Acc_RazSoc 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   420
            Width           =   3315
         End
         Begin VB.TextBox txt_Acc_NumFax 
            Height          =   315
            Left            =   8100
            MaxLength       =   12
            TabIndex        =   55
            Text            =   "Text1"
            Top             =   3060
            Width           =   1640
         End
         Begin VB.TextBox txt_Acc_Refere 
            Height          =   315
            Left            =   8100
            MaxLength       =   250
            TabIndex        =   52
            Text            =   "Text1"
            Top             =   2730
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Acc_PrvDir 
            Height          =   315
            Left            =   8100
            TabIndex        =   50
            Text            =   "cmb_PrvDir"
            Top             =   2400
            Width           =   3315
         End
         Begin VB.TextBox txt_Acc_NomZon 
            Height          =   315
            Left            =   8100
            MaxLength       =   120
            TabIndex        =   48
            Text            =   "Text1"
            Top             =   2070
            Width           =   3315
         End
         Begin VB.TextBox txt_Acc_Interi 
            Height          =   315
            Left            =   9750
            MaxLength       =   15
            TabIndex        =   46
            Text            =   "Text1"
            Top             =   1740
            Width           =   1640
         End
         Begin VB.TextBox txt_Acc_Numero 
            Height          =   315
            Left            =   8100
            MaxLength       =   15
            TabIndex        =   45
            Text            =   "Text1"
            Top             =   1740
            Width           =   1640
         End
         Begin VB.ComboBox cmb_Acc_TipVia 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1410
            Width           =   3315
         End
         Begin VB.TextBox txt_Acc_Telef2 
            Height          =   315
            Left            =   3630
            MaxLength       =   12
            TabIndex        =   54
            Text            =   "Text1"
            Top             =   3060
            Width           =   1640
         End
         Begin VB.TextBox txt_Acc_NomCom 
            Height          =   315
            Left            =   8100
            MaxLength       =   250
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   420
            Width           =   3315
         End
         Begin Threed.SSPanel pnl_Acc_FlgEmp 
            Height          =   315
            Left            =   10980
            TabIndex        =   251
            Top             =   90
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
         Begin Threed.SSPanel SSPanel17 
            Height          =   90
            Left            =   30
            TabIndex        =   252
            Top             =   3390
            Width           =   11445
            _Version        =   65536
            _ExtentX        =   20188
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
         Begin EditLib.fpDoubleSingle ipp_Acc_IngNet 
            Height          =   315
            Left            =   1980
            TabIndex        =   56
            Top             =   3540
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
         Begin EditLib.fpDateTime ipp_Acc_FecAnt 
            Height          =   315
            Left            =   1980
            TabIndex        =   58
            Top             =   4200
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
         Begin EditLib.fpDoubleSingle ipp_Acc_PorAcc 
            Height          =   315
            Left            =   1980
            TabIndex        =   57
            Top             =   3870
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
         Begin VB.Label lbl_General 
            Caption         =   "Giro Comercial:"
            Height          =   285
            Index           =   35
            Left            =   60
            TabIndex        =   272
            Top             =   750
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Razón Social:"
            Height          =   285
            Index           =   34
            Left            =   60
            TabIndex        =   271
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Index           =   33
            Left            =   6090
            TabIndex        =   270
            Top             =   90
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   32
            Left            =   60
            TabIndex        =   269
            Top             =   90
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono:"
            Height          =   285
            Index           =   31
            Left            =   60
            TabIndex        =   268
            Top             =   3060
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Distrito:"
            Height          =   315
            Index           =   30
            Left            =   60
            TabIndex        =   267
            Top             =   2730
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Departamento:"
            Height          =   315
            Index           =   29
            Left            =   60
            TabIndex        =   266
            Top             =   2400
            Width           =   1425
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Index           =   28
            Left            =   60
            TabIndex        =   265
            Top             =   2070
            Width           =   1455
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Index           =   27
            Left            =   60
            TabIndex        =   264
            Top             =   1740
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fax:"
            Height          =   285
            Index           =   25
            Left            =   6090
            TabIndex        =   263
            Top             =   3060
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Referencia:"
            Height          =   285
            Index           =   24
            Left            =   6090
            TabIndex        =   262
            Top             =   2730
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Provincia:"
            Height          =   315
            Index           =   23
            Left            =   6090
            TabIndex        =   261
            Top             =   2400
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Index           =   22
            Left            =   6090
            TabIndex        =   260
            Top             =   2070
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Index           =   21
            Left            =   6090
            TabIndex        =   259
            Top             =   1740
            Width           =   1935
         End
         Begin VB.Label lbl_General 
            Caption         =   "Giro Comercial (Especif.):"
            Height          =   285
            Index           =   20
            Left            =   60
            TabIndex        =   258
            Top             =   1080
            Width           =   1875
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Vía:"
            Height          =   285
            Index           =   19
            Left            =   60
            TabIndex        =   257
            Top             =   1410
            Width           =   1545
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Index           =   18
            Left            =   6090
            TabIndex        =   256
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label Label69 
            Caption         =   "Ingresos (S/.):"
            Height          =   285
            Left            =   60
            TabIndex        =   255
            Top             =   3540
            Width           =   1965
         End
         Begin VB.Label Label68 
            Caption         =   "Porc. Accionariado:"
            Height          =   285
            Left            =   60
            TabIndex        =   254
            Top             =   3870
            Width           =   1785
         End
         Begin VB.Label Label67 
            Caption         =   "Fecha Antigüedad:"
            Height          =   315
            Left            =   60
            TabIndex        =   253
            Top             =   4200
            Width           =   1635
         End
      End
      Begin Threed.SSPanel pnl_TraRen 
         Height          =   5895
         Left            =   30
         TabIndex        =   273
         Top             =   2700
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   10398
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
         Begin VB.ComboBox cmb_Ren_GirCom 
            Height          =   315
            Left            =   1980
            TabIndex        =   61
            Text            =   "cmb_GirCom"
            Top             =   420
            Width           =   9465
         End
         Begin VB.CheckBox chk_Alqui3 
            Caption         =   "3."
            Height          =   255
            Left            =   60
            TabIndex        =   75
            Top             =   2010
            Width           =   435
         End
         Begin VB.CheckBox chk_Alqui2 
            Caption         =   "2."
            Height          =   255
            Left            =   60
            TabIndex        =   68
            Top             =   1680
            Width           =   435
         End
         Begin VB.TextBox txt_Ren_NomAr3 
            Height          =   315
            Left            =   3390
            MaxLength       =   250
            TabIndex        =   77
            Text            =   "Text1"
            Top             =   1980
            Width           =   2775
         End
         Begin VB.TextBox txt_Ren_Direc3 
            Height          =   315
            Left            =   600
            MaxLength       =   250
            TabIndex        =   76
            Text            =   "Text1"
            Top             =   1980
            Width           =   2775
         End
         Begin VB.TextBox txt_Ren_Tele23 
            Height          =   315
            Left            =   7470
            MaxLength       =   12
            TabIndex        =   79
            Text            =   "Text1"
            Top             =   1980
            Width           =   1275
         End
         Begin VB.TextBox txt_Ren_Tele13 
            Height          =   315
            Left            =   6180
            MaxLength       =   12
            TabIndex        =   78
            Text            =   "Text1"
            Top             =   1980
            Width           =   1275
         End
         Begin VB.TextBox txt_Ren_NomAr2 
            Height          =   315
            Left            =   3390
            MaxLength       =   250
            TabIndex        =   70
            Text            =   "Text1"
            Top             =   1650
            Width           =   2775
         End
         Begin VB.TextBox txt_Ren_Direc2 
            Height          =   315
            Left            =   600
            MaxLength       =   250
            TabIndex        =   69
            Text            =   "Text1"
            Top             =   1650
            Width           =   2775
         End
         Begin VB.TextBox txt_Ren_Tele22 
            Height          =   315
            Left            =   7470
            MaxLength       =   12
            TabIndex        =   72
            Text            =   "Text1"
            Top             =   1650
            Width           =   1275
         End
         Begin VB.TextBox txt_Ren_Tele12 
            Height          =   315
            Left            =   6180
            MaxLength       =   12
            TabIndex        =   71
            Text            =   "Text1"
            Top             =   1650
            Width           =   1275
         End
         Begin VB.TextBox txt_Ren_NomAr1 
            Height          =   315
            Left            =   3390
            MaxLength       =   250
            TabIndex        =   63
            Text            =   "Text1"
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox txt_Ren_NumDoc 
            Height          =   315
            Left            =   8100
            MaxLength       =   11
            TabIndex        =   60
            Text            =   "Text1"
            Top             =   90
            Width           =   2355
         End
         Begin VB.ComboBox cmb_Ren_TipDoc 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   90
            Width           =   3315
         End
         Begin VB.TextBox txt_Ren_Direc1 
            Height          =   315
            Left            =   600
            MaxLength       =   250
            TabIndex        =   62
            Text            =   "Text1"
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox txt_Ren_Tele21 
            Height          =   315
            Left            =   7470
            MaxLength       =   12
            TabIndex        =   65
            Text            =   "Text1"
            Top             =   1320
            Width           =   1275
         End
         Begin VB.TextBox txt_Ren_Tele11 
            Height          =   315
            Left            =   6180
            MaxLength       =   12
            TabIndex        =   64
            Text            =   "Text1"
            Top             =   1320
            Width           =   1275
         End
         Begin Threed.SSPanel SSPanel19 
            Height          =   90
            Left            =   30
            TabIndex        =   274
            Top             =   840
            Width           =   11445
            _Version        =   65536
            _ExtentX        =   20188
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
         Begin EditLib.fpDoubleSingle ipp_Ren_AlqMe1 
            Height          =   315
            Left            =   8760
            TabIndex        =   66
            Top             =   1320
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
            MaxValue        =   "999999"
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
         Begin EditLib.fpDateTime ipp_Ren_FIAlq1 
            Height          =   315
            Left            =   10110
            TabIndex        =   67
            Top             =   1320
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
         Begin EditLib.fpDoubleSingle ipp_Ren_AlqMe2 
            Height          =   315
            Left            =   8760
            TabIndex        =   73
            Top             =   1650
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
            MaxValue        =   "999999"
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
         Begin EditLib.fpDateTime ipp_Ren_FIAlq2 
            Height          =   315
            Left            =   10110
            TabIndex        =   74
            Top             =   1650
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
         Begin EditLib.fpDoubleSingle ipp_Ren_AlqMe3 
            Height          =   315
            Left            =   8760
            TabIndex        =   80
            Top             =   1980
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
            MaxValue        =   "999999"
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
         Begin EditLib.fpDateTime ipp_Ren_FIAlq3 
            Height          =   315
            Left            =   10110
            TabIndex        =   81
            Top             =   1980
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
         Begin EditLib.fpDoubleSingle ipp_Ren_IngNet 
            Height          =   315
            Left            =   1920
            TabIndex        =   82
            Top             =   2520
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
            MaxValue        =   "9999999"
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   90
            Left            =   30
            TabIndex        =   284
            Top             =   2370
            Width           =   11445
            _Version        =   65536
            _ExtentX        =   20188
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
         Begin VB.Label lbl_General 
            Caption         =   "Giro Comercial:"
            Height          =   285
            Index           =   68
            Left            =   60
            TabIndex        =   289
            Top             =   420
            Width           =   1395
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "1."
            Height          =   285
            Left            =   60
            TabIndex        =   285
            Top             =   1350
            Width           =   405
         End
         Begin VB.Label Label35 
            Caption         =   "Ingresos (S/.):"
            Height          =   285
            Left            =   60
            TabIndex        =   283
            Top             =   2520
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "F. Inicio Alq."
            Height          =   285
            Left            =   10110
            TabIndex        =   282
            Top             =   1020
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Alq. Mens. US$"
            Height          =   285
            Left            =   8760
            TabIndex        =   281
            Top             =   1020
            Width           =   1275
         End
         Begin VB.Label Label9 
            Caption         =   "Telef. 2"
            Height          =   285
            Left            =   7470
            TabIndex        =   280
            Top             =   1020
            Width           =   705
         End
         Begin VB.Label Label8 
            Caption         =   "Telef. 1"
            Height          =   285
            Left            =   6180
            TabIndex        =   279
            Top             =   1020
            Width           =   705
         End
         Begin VB.Label Label6 
            Caption         =   "Nombre Arrendatario"
            Height          =   285
            Left            =   3390
            TabIndex        =   278
            Top             =   1020
            Width           =   1965
         End
         Begin VB.Label lbl_General 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Index           =   82
            Left            =   6090
            TabIndex        =   277
            Top             =   90
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   81
            Left            =   60
            TabIndex        =   276
            Top             =   90
            Width           =   1635
         End
         Begin VB.Label Label3 
            Caption         =   "Dirección Propiedad"
            Height          =   285
            Left            =   600
            TabIndex        =   275
            Top             =   1020
            Width           =   1965
         End
      End
      Begin Threed.SSPanel pnl_TraCom 
         Height          =   5895
         Left            =   30
         TabIndex        =   222
         Top             =   2700
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   10398
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
         Begin VB.CommandButton cmd_Com_BusEmp 
            Caption         =   "..."
            Height          =   315
            Left            =   10530
            TabIndex        =   85
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   90
            Width           =   435
         End
         Begin VB.ComboBox cmb_Com_RegTri 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   4200
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Com_TipLoc 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   4560
            Width           =   3315
         End
         Begin VB.TextBox txt_Com_Tl1Arr 
            Height          =   315
            Left            =   8100
            MaxLength       =   12
            TabIndex        =   111
            Text            =   "Text1"
            Top             =   4890
            Width           =   1640
         End
         Begin VB.TextBox txt_Com_Tl2Arr 
            Height          =   315
            Left            =   9750
            MaxLength       =   12
            TabIndex        =   112
            Text            =   "Text1"
            Top             =   4890
            Width           =   1640
         End
         Begin VB.TextBox txt_Com_NomArr 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   110
            Text            =   "Text1"
            Top             =   4890
            Width           =   3315
         End
         Begin VB.TextBox txt_Com_NomCom 
            Height          =   315
            Left            =   8100
            MaxLength       =   250
            TabIndex        =   87
            Text            =   "Text1"
            Top             =   420
            Width           =   3315
         End
         Begin VB.TextBox txt_Com_Telef2 
            Height          =   315
            Left            =   3630
            MaxLength       =   12
            TabIndex        =   101
            Text            =   "Text1"
            Top             =   3060
            Width           =   1640
         End
         Begin VB.ComboBox cmb_Com_TipVia 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   1410
            Width           =   3315
         End
         Begin VB.TextBox txt_Com_Numero 
            Height          =   315
            Left            =   8100
            MaxLength       =   15
            TabIndex        =   92
            Text            =   "Text1"
            Top             =   1740
            Width           =   1640
         End
         Begin VB.TextBox txt_Com_Interi 
            Height          =   315
            Left            =   9750
            MaxLength       =   15
            TabIndex        =   93
            Text            =   "Text1"
            Top             =   1740
            Width           =   1640
         End
         Begin VB.TextBox txt_Com_NomZon 
            Height          =   315
            Left            =   8100
            MaxLength       =   120
            TabIndex        =   95
            Text            =   "Text1"
            Top             =   2070
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Com_PrvDir 
            Height          =   315
            Left            =   8100
            TabIndex        =   97
            Text            =   "cmb_PrvDir"
            Top             =   2400
            Width           =   3315
         End
         Begin VB.TextBox txt_Com_Refere 
            Height          =   315
            Left            =   8100
            MaxLength       =   250
            TabIndex        =   99
            Text            =   "Text1"
            Top             =   2730
            Width           =   3315
         End
         Begin VB.TextBox txt_Com_NumFax 
            Height          =   315
            Left            =   8100
            MaxLength       =   12
            TabIndex        =   102
            Text            =   "Text1"
            Top             =   3060
            Width           =   1640
         End
         Begin VB.TextBox txt_Com_RazSoc 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   86
            Text            =   "Text1"
            Top             =   420
            Width           =   3315
         End
         Begin VB.TextBox txt_Com_GirCom 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   89
            Text            =   "Text1"
            Top             =   1080
            Width           =   9465
         End
         Begin VB.ComboBox cmb_Com_TipDoc 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   90
            Width           =   3315
         End
         Begin VB.TextBox txt_Com_NumDoc 
            Height          =   315
            Left            =   8100
            MaxLength       =   11
            TabIndex        =   84
            Text            =   "Text1"
            Top             =   90
            Width           =   2355
         End
         Begin VB.ComboBox cmb_Com_GirCom 
            Height          =   315
            Left            =   1980
            TabIndex        =   88
            Text            =   "cmb_GirCom"
            Top             =   750
            Width           =   9465
         End
         Begin VB.TextBox txt_Com_NomVia 
            Height          =   315
            Left            =   1980
            MaxLength       =   120
            TabIndex        =   91
            Text            =   "Text1"
            Top             =   1740
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Com_TipZon 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   2070
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Com_DptDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   96
            Text            =   "cmb_DptDir"
            Top             =   2400
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Com_DstDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   98
            Text            =   "cmb_DstDir"
            Top             =   2730
            Width           =   3315
         End
         Begin VB.TextBox txt_Com_Telef1 
            Height          =   315
            Left            =   1980
            MaxLength       =   12
            TabIndex        =   100
            Text            =   "Text1"
            Top             =   3060
            Width           =   1640
         End
         Begin Threed.SSPanel SSPanel15 
            Height          =   90
            Left            =   30
            TabIndex        =   223
            Top             =   3390
            Width           =   11445
            _Version        =   65536
            _ExtentX        =   20188
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
         Begin EditLib.fpDoubleSingle ipp_Com_IngNet 
            Height          =   315
            Left            =   1980
            TabIndex        =   103
            Top             =   3540
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
         Begin EditLib.fpDateTime ipp_Com_FecIni 
            Height          =   315
            Left            =   8100
            TabIndex        =   105
            Top             =   3900
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
         Begin EditLib.fpDoubleSingle ipp_Com_VtaMen 
            Height          =   315
            Left            =   1980
            TabIndex        =   104
            Top             =   3870
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
         Begin EditLib.fpDoubleSingle ipp_Com_AlqMen 
            Height          =   315
            Left            =   8100
            TabIndex        =   109
            Top             =   4560
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
         Begin EditLib.fpDoubleSingle ipp_Com_PorPar 
            Height          =   315
            Left            =   8100
            TabIndex        =   107
            Top             =   4230
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
            MaxValue        =   "100"
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
         Begin Threed.SSPanel pnl_Com_FlgEmp 
            Height          =   315
            Left            =   11010
            TabIndex        =   288
            Top             =   90
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
         Begin VB.Label Label66 
            Caption         =   "Régimen Tributario:"
            Height          =   315
            Left            =   60
            TabIndex        =   249
            Top             =   4200
            Width           =   1635
         End
         Begin VB.Label Label65 
            Caption         =   "Ventas Mensuales S/."
            Height          =   285
            Left            =   60
            TabIndex        =   248
            Top             =   3870
            Width           =   1785
         End
         Begin VB.Label Label64 
            Caption         =   "Ingresos (S/.):"
            Height          =   285
            Left            =   60
            TabIndex        =   247
            Top             =   3540
            Width           =   1965
         End
         Begin VB.Label Label47 
            Caption         =   "Fecha de Inicio Operac.:"
            Height          =   315
            Left            =   6090
            TabIndex        =   246
            Top             =   3900
            Width           =   1905
         End
         Begin VB.Label Label45 
            Caption         =   "Tipo Local Comercial:"
            Height          =   315
            Left            =   60
            TabIndex        =   245
            Top             =   4560
            Width           =   1605
         End
         Begin VB.Label Label44 
            Caption         =   "Teléfonos:"
            Height          =   285
            Left            =   6090
            TabIndex        =   244
            Top             =   4890
            Width           =   1665
         End
         Begin VB.Label Label43 
            Caption         =   "Nombre Arrendador:"
            Height          =   285
            Left            =   60
            TabIndex        =   243
            Top             =   4890
            Width           =   1485
         End
         Begin VB.Label Label42 
            Caption         =   "Monto Alq. Mens. US$:"
            Height          =   285
            Left            =   6090
            TabIndex        =   242
            Top             =   4560
            Width           =   1785
         End
         Begin VB.Label Label40 
            Caption         =   "% Participación:"
            Height          =   285
            Left            =   6090
            TabIndex        =   241
            Top             =   4230
            Width           =   1785
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Index           =   11
            Left            =   6090
            TabIndex        =   240
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Vía:"
            Height          =   285
            Index           =   4
            Left            =   60
            TabIndex        =   239
            Top             =   1410
            Width           =   1545
         End
         Begin VB.Label lbl_General 
            Caption         =   "Giro Comercial (Especif.):"
            Height          =   285
            Index           =   12
            Left            =   60
            TabIndex        =   238
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Index           =   13
            Left            =   6090
            TabIndex        =   237
            Top             =   1740
            Width           =   1935
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Index           =   14
            Left            =   6090
            TabIndex        =   236
            Top             =   2070
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Provincia:"
            Height          =   315
            Index           =   15
            Left            =   6090
            TabIndex        =   235
            Top             =   2400
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Referencia:"
            Height          =   285
            Index           =   16
            Left            =   6090
            TabIndex        =   234
            Top             =   2730
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fax:"
            Height          =   285
            Index           =   17
            Left            =   6090
            TabIndex        =   233
            Top             =   3060
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Index           =   5
            Left            =   60
            TabIndex        =   232
            Top             =   1740
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Index           =   6
            Left            =   60
            TabIndex        =   231
            Top             =   2070
            Width           =   1455
         End
         Begin VB.Label lbl_General 
            Caption         =   "Departamento:"
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   230
            Top             =   2400
            Width           =   1425
         End
         Begin VB.Label lbl_General 
            Caption         =   "Distrito:"
            Height          =   315
            Index           =   8
            Left            =   60
            TabIndex        =   229
            Top             =   2730
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono:"
            Height          =   285
            Index           =   9
            Left            =   60
            TabIndex        =   228
            Top             =   3060
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   227
            Top             =   90
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Index           =   10
            Left            =   6090
            TabIndex        =   226
            Top             =   90
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Razón Social:"
            Height          =   285
            Index           =   1
            Left            =   60
            TabIndex        =   225
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Giro Comercial:"
            Height          =   285
            Index           =   3
            Left            =   60
            TabIndex        =   224
            Top             =   750
            Width           =   1365
         End
      End
      Begin Threed.SSPanel pnl_TraDep 
         Height          =   5895
         Left            =   30
         TabIndex        =   162
         Top             =   2700
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   10398
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
         Begin VB.CheckBox chk_Dep_Sucurs 
            Height          =   285
            Left            =   1980
            TabIndex        =   120
            Top             =   1410
            Width           =   285
         End
         Begin VB.CommandButton cmd_Dep_BusEmp 
            Caption         =   "..."
            Height          =   315
            Left            =   10500
            TabIndex        =   115
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   90
            Width           =   435
         End
         Begin VB.ComboBox cmb_Dep_FreHab 
            Height          =   315
            Left            =   8100
            Style           =   2  'Dropdown List
            TabIndex        =   138
            Top             =   4200
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Dep_NomCar 
            Height          =   315
            Left            =   1980
            TabIndex        =   139
            Text            =   "cmb_Dep_NomCar"
            Top             =   4530
            Width           =   3315
         End
         Begin VB.TextBox txt_Dep_NomCar 
            Height          =   315
            Left            =   8100
            MaxLength       =   250
            TabIndex        =   140
            Text            =   "Text1"
            Top             =   4530
            Width           =   3315
         End
         Begin VB.TextBox txt_Dep_TelDir 
            Height          =   315
            Left            =   8100
            MaxLength       =   12
            TabIndex        =   145
            Text            =   "Text1"
            Top             =   5190
            Width           =   1640
         End
         Begin VB.TextBox txt_Dep_NumAnx 
            Height          =   315
            Left            =   1980
            MaxLength       =   5
            TabIndex        =   143
            Text            =   "Text1"
            Top             =   5190
            Width           =   1640
         End
         Begin VB.TextBox txt_Dep_DirEle 
            Height          =   315
            Left            =   1980
            MaxLength       =   120
            TabIndex        =   149
            Text            =   "Text1"
            Top             =   5520
            Width           =   3315
         End
         Begin VB.TextBox txt_Dep_NomAre 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   141
            Text            =   "Text1"
            Top             =   4860
            Width           =   3315
         End
         Begin VB.TextBox txt_Dep_Celula 
            Height          =   315
            Left            =   9780
            MaxLength       =   12
            TabIndex        =   147
            Text            =   "Text1"
            Top             =   5190
            Width           =   1640
         End
         Begin VB.TextBox txt_Dep_TeleRH 
            Height          =   315
            Left            =   1980
            MaxLength       =   12
            TabIndex        =   135
            Text            =   "Text1"
            Top             =   3720
            Width           =   1640
         End
         Begin VB.TextBox txt_Dep_AnexRH 
            Height          =   315
            Left            =   3630
            MaxLength       =   12
            TabIndex        =   136
            Text            =   "Text1"
            Top             =   3720
            Width           =   1640
         End
         Begin VB.TextBox txt_Dep_Telef1 
            Height          =   315
            Left            =   1980
            MaxLength       =   12
            TabIndex        =   132
            Text            =   "Text1"
            Top             =   3390
            Width           =   1640
         End
         Begin VB.ComboBox cmb_Dep_DstDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   130
            Text            =   "cmb_DstDir"
            Top             =   3060
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Dep_DptDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   128
            Text            =   "cmb_DptDir"
            Top             =   2730
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Dep_TipZon 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   2400
            Width           =   3315
         End
         Begin VB.TextBox txt_Dep_NomVia 
            Height          =   315
            Left            =   1980
            MaxLength       =   120
            TabIndex        =   123
            Text            =   "Text1"
            Top             =   2070
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Dep_GirCom 
            Height          =   315
            Left            =   1980
            TabIndex        =   118
            Text            =   "cmb_GirCom"
            Top             =   750
            Width           =   9465
         End
         Begin VB.TextBox txt_Dep_NumDoc 
            Height          =   315
            Left            =   8100
            MaxLength       =   11
            TabIndex        =   114
            Text            =   "Text1"
            Top             =   90
            Width           =   2355
         End
         Begin VB.ComboBox cmb_Dep_TipDoc 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   113
            Top             =   90
            Width           =   3315
         End
         Begin VB.TextBox txt_Dep_GirCom 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   119
            Text            =   "Text1"
            Top             =   1080
            Width           =   9465
         End
         Begin VB.TextBox txt_Dep_RazSoc 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   116
            Text            =   "Text1"
            Top             =   420
            Width           =   3315
         End
         Begin VB.TextBox txt_Dep_NumFax 
            Height          =   315
            Left            =   8100
            MaxLength       =   12
            TabIndex        =   134
            Text            =   "Text1"
            Top             =   3390
            Width           =   1640
         End
         Begin VB.TextBox txt_Dep_Refere 
            Height          =   315
            Left            =   8100
            MaxLength       =   250
            TabIndex        =   131
            Text            =   "Text1"
            Top             =   3060
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Dep_PrvDir 
            Height          =   315
            Left            =   8100
            TabIndex        =   129
            Text            =   "cmb_PrvDir"
            Top             =   2730
            Width           =   3315
         End
         Begin VB.TextBox txt_Dep_NomZon 
            Height          =   315
            Left            =   8100
            MaxLength       =   120
            TabIndex        =   127
            Text            =   "Text1"
            Top             =   2400
            Width           =   3315
         End
         Begin VB.TextBox txt_Dep_Interi 
            Height          =   315
            Left            =   9750
            MaxLength       =   15
            TabIndex        =   125
            Text            =   "Text1"
            Top             =   2070
            Width           =   1640
         End
         Begin VB.TextBox txt_Dep_Numero 
            Height          =   315
            Left            =   8100
            MaxLength       =   15
            TabIndex        =   124
            Text            =   "Text1"
            Top             =   2070
            Width           =   1640
         End
         Begin VB.TextBox txt_Dep_Sucurs 
            Height          =   315
            Left            =   2250
            MaxLength       =   250
            TabIndex        =   121
            Text            =   "Text1"
            Top             =   1410
            Width           =   3045
         End
         Begin VB.ComboBox cmb_Dep_TipVia 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   1740
            Width           =   3315
         End
         Begin VB.TextBox txt_Dep_Telef2 
            Height          =   315
            Left            =   3630
            MaxLength       =   12
            TabIndex        =   133
            Text            =   "Text1"
            Top             =   3390
            Width           =   1640
         End
         Begin VB.TextBox txt_Dep_NomCom 
            Height          =   315
            Left            =   8130
            MaxLength       =   250
            TabIndex        =   117
            Text            =   "Text1"
            Top             =   420
            Width           =   3315
         End
         Begin Threed.SSPanel pnl_Dep_FlgEmp 
            Height          =   315
            Left            =   10980
            TabIndex        =   163
            Top             =   90
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   90
            Left            =   30
            TabIndex        =   183
            Top             =   4080
            Width           =   11445
            _Version        =   65536
            _ExtentX        =   20188
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
         Begin EditLib.fpDoubleSingle ipp_Dep_IngNet 
            Height          =   315
            Left            =   1980
            TabIndex        =   137
            Top             =   4200
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
         Begin EditLib.fpDateTime ipp_Dep_FecIng 
            Height          =   315
            Left            =   8100
            TabIndex        =   142
            Top             =   4860
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
         Begin EditLib.fpDateTime ipp_Dep_FecCes 
            Height          =   315
            Left            =   8100
            TabIndex        =   151
            Top             =   5520
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
            Caption         =   "Telf. Directo / Celular:"
            Height          =   285
            Index           =   59
            Left            =   6090
            TabIndex        =   193
            Top             =   5160
            Width           =   1575
         End
         Begin VB.Label lbl_General 
            Caption         =   "Ingresos (S/.):"
            Height          =   285
            Index           =   61
            Left            =   60
            TabIndex        =   192
            Top             =   4200
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Frecuencia Haberes:"
            Height          =   315
            Index           =   56
            Left            =   6090
            TabIndex        =   191
            Top             =   4200
            Width           =   1905
         End
         Begin VB.Label lbl_General 
            Caption         =   "Cargo:"
            Height          =   285
            Index           =   62
            Left            =   60
            TabIndex        =   190
            Top             =   4530
            Width           =   975
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Index           =   58
            Left            =   6090
            TabIndex        =   189
            Top             =   4860
            Width           =   1905
         End
         Begin VB.Label lbl_General 
            Caption         =   "Anexo:"
            Height          =   285
            Index           =   64
            Left            =   60
            TabIndex        =   188
            Top             =   5190
            Width           =   1575
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha de Cese:"
            Height          =   315
            Index           =   66
            Left            =   6090
            TabIndex        =   187
            Top             =   5520
            Width           =   1905
         End
         Begin VB.Label lbl_General 
            Caption         =   "Cargo (Especificar):"
            Height          =   285
            Index           =   57
            Left            =   6090
            TabIndex        =   186
            Top             =   4530
            Width           =   2055
         End
         Begin VB.Label lbl_General 
            Caption         =   "E-mail:"
            Height          =   285
            Index           =   60
            Left            =   60
            TabIndex        =   185
            Top             =   5520
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Area:"
            Height          =   285
            Index           =   63
            Left            =   60
            TabIndex        =   184
            Top             =   4860
            Width           =   1605
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono/Anexo RR.HH:"
            Height          =   285
            Index           =   47
            Left            =   60
            TabIndex        =   182
            Top             =   3720
            Width           =   1815
         End
         Begin VB.Label lbl_General 
            Caption         =   "Giro Comercial:"
            Height          =   285
            Index           =   39
            Left            =   60
            TabIndex        =   181
            Top             =   750
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Razón Social:"
            Height          =   285
            Index           =   37
            Left            =   60
            TabIndex        =   180
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Index           =   48
            Left            =   6090
            TabIndex        =   179
            Top             =   90
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   36
            Left            =   60
            TabIndex        =   178
            Top             =   90
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono:"
            Height          =   285
            Index           =   46
            Left            =   60
            TabIndex        =   177
            Top             =   3390
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Distrito:"
            Height          =   315
            Index           =   45
            Left            =   60
            TabIndex        =   176
            Top             =   3060
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Departamento:"
            Height          =   315
            Index           =   44
            Left            =   60
            TabIndex        =   175
            Top             =   2730
            Width           =   1425
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Index           =   43
            Left            =   60
            TabIndex        =   174
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Index           =   42
            Left            =   60
            TabIndex        =   173
            Top             =   2070
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fax:"
            Height          =   285
            Index           =   55
            Left            =   6090
            TabIndex        =   172
            Top             =   3390
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Referencia:"
            Height          =   285
            Index           =   54
            Left            =   6090
            TabIndex        =   171
            Top             =   3060
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Provincia:"
            Height          =   315
            Index           =   53
            Left            =   6090
            TabIndex        =   170
            Top             =   2730
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Index           =   52
            Left            =   6090
            TabIndex        =   169
            Top             =   2400
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Index           =   51
            Left            =   6090
            TabIndex        =   168
            Top             =   2070
            Width           =   1935
         End
         Begin VB.Label lbl_General 
            Caption         =   "Sucursal:"
            Height          =   285
            Index           =   40
            Left            =   60
            TabIndex        =   167
            Top             =   1410
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Giro Comercial (Especif.):"
            Height          =   285
            Index           =   50
            Left            =   60
            TabIndex        =   166
            Top             =   1080
            Width           =   1875
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Vía:"
            Height          =   285
            Index           =   41
            Left            =   60
            TabIndex        =   165
            Top             =   1740
            Width           =   1545
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Index           =   49
            Left            =   6090
            TabIndex        =   164
            Top             =   420
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "frm_IngSol_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Dep_GirCom()  As moddat_tpo_Genera
Dim l_arr_Ind_GirCom()  As moddat_tpo_Genera
Dim l_arr_Com_GirCom()  As moddat_tpo_Genera
Dim l_arr_Acc_GirCom()  As moddat_tpo_Genera
Dim l_arr_Ren_GirCom()  As moddat_tpo_Genera

Dim l_arr_Dep_NomCar()  As moddat_tpo_Genera
Dim l_arr_Ind_NomCar()  As moddat_tpo_Genera

Dim l_int_FlgCmb        As Integer
Dim l_str_Dep_DstDir    As String
Dim l_str_Dep_DptDir    As String
Dim l_str_Dep_PrvDir    As String
Dim l_str_Ind_DstDir    As String
Dim l_str_Ind_DptDir    As String
Dim l_str_Ind_PrvDir    As String
Dim l_str_Com_DstDir    As String
Dim l_str_Com_DptDir    As String
Dim l_str_Com_PrvDir    As String
Dim l_str_Acc_DstDir    As String
Dim l_str_Acc_DptDir    As String
Dim l_str_Acc_PrvDir    As String
Dim l_str_Dep_GirCom    As String
Dim l_str_Dep_NomCar    As String
Dim l_str_Ind_GirCom    As String
Dim l_str_Ind_NomCar    As String
Dim l_str_Acc_GirCom    As String
Dim l_str_Com_GirCom    As String
Dim l_str_Ren_GirCom    As String
Dim l_int_FlgGrb        As Integer

Private Sub chk_Alqui2_Click()
   If chk_Alqui2.Value = 0 Then
      txt_Ren_Direc2.Text = ""
      txt_Ren_NomAr2.Text = ""
      txt_Ren_Tele12.Text = ""
      txt_Ren_Tele22.Text = ""
      ipp_Ren_AlqMe2.Value = 0
      ipp_Ren_FIAlq2.Text = Format(CDate(CDate(moddat_g_str_FecSis) - CDate(365)), "dd/mm/yyyy")

      txt_Ren_Direc2.Enabled = False
      txt_Ren_NomAr2.Enabled = False
      txt_Ren_Tele12.Enabled = False
      txt_Ren_Tele22.Enabled = False
      ipp_Ren_AlqMe2.Enabled = False
      ipp_Ren_FIAlq2.Enabled = False
   ElseIf chk_Alqui2.Value = 1 Then
      txt_Ren_Direc2.Enabled = True
      txt_Ren_NomAr2.Enabled = True
      txt_Ren_Tele12.Enabled = True
      txt_Ren_Tele22.Enabled = True
      ipp_Ren_AlqMe2.Enabled = True
      ipp_Ren_FIAlq2.Enabled = True
      
      Call gs_SetFocus(txt_Ren_Direc2)
   End If
End Sub

Private Sub chk_Alqui3_Click()
   If chk_Alqui2.Value = 0 Then
      chk_Alqui3.Value = 0
      Exit Sub
   End If

   If chk_Alqui3.Value = 0 Then
      txt_Ren_Direc3.Text = ""
      txt_Ren_NomAr3.Text = ""
      txt_Ren_Tele13.Text = ""
      txt_Ren_Tele23.Text = ""
      ipp_Ren_AlqMe3.Value = 0
      ipp_Ren_FIAlq3.Text = Format(CDate(CDate(moddat_g_str_FecSis) - CDate(365)), "dd/mm/yyyy")

      txt_Ren_Direc3.Enabled = False
      txt_Ren_NomAr3.Enabled = False
      txt_Ren_Tele13.Enabled = False
      txt_Ren_Tele23.Enabled = False
      ipp_Ren_AlqMe3.Enabled = False
      ipp_Ren_FIAlq3.Enabled = False
   ElseIf chk_Alqui3.Value = 1 Then
      txt_Ren_Direc3.Enabled = True
      txt_Ren_NomAr3.Enabled = True
      txt_Ren_Tele13.Enabled = True
      txt_Ren_Tele23.Enabled = True
      ipp_Ren_AlqMe3.Enabled = True
      ipp_Ren_FIAlq3.Enabled = True
      
      Call gs_SetFocus(txt_Ren_Direc3)
   End If
End Sub

Private Sub chk_Dep_Sucurs_Click()
   If chk_Dep_Sucurs.Value = 1 Then
      txt_Dep_Sucurs.Enabled = True
      
      Call gs_SetFocus(txt_Dep_Sucurs)
      
      cmb_Dep_TipVia.Enabled = True
      txt_Dep_NomVia.Enabled = True
      txt_Dep_Numero.Enabled = True
      txt_Dep_Interi.Enabled = True
      cmb_Dep_TipZon.Enabled = True
      txt_Dep_NomZon.Enabled = True
      cmb_Dep_DptDir.Enabled = True
      cmb_Dep_PrvDir.Enabled = True
      cmb_Dep_DstDir.Enabled = True
      txt_Dep_Refere.Enabled = True
      txt_Dep_Telef1.Enabled = True
      txt_Dep_Telef2.Enabled = True
      txt_Dep_NumFax.Enabled = True
      txt_Dep_TeleRH.Enabled = True
      txt_Dep_AnexRH.Enabled = True
   Else
      txt_Dep_Sucurs.Text = ""
      
      If cmb_Dep_TipDoc.ListIndex = -1 Then
         Exit Sub
      End If
      
      If Len(Trim(txt_Dep_NumDoc.Text)) = 0 Then
         Exit Sub
      End If
      
      If Not gf_Valida_RUC(Mid(txt_Dep_NumDoc.Text, 1, Len(txt_Dep_NumDoc.Text) - 1), Right(txt_Dep_NumDoc.Text, 1)) Then
         Exit Sub
      End If
      
      If pnl_Dep_FlgEmp.Caption <> "NR" Then
         Call fs_BusEmp(cmb_Dep_TipDoc.ItemData(cmb_Dep_TipDoc.ListIndex), txt_Dep_NumDoc)
      End If
   End If
End Sub

Private Sub chk_Dep_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_Dep_TipVia.Enabled Then
         Call gs_SetFocus(cmb_Dep_TipVia)
      Else
         Call gs_SetFocus(ipp_Dep_IngNet)
      End If
   End If
End Sub

Private Sub cmb_ActEco_Click()
   If cmb_ActEco.ListIndex <> -1 Then
      Select Case cmb_ActEco.ItemData(cmb_ActEco.ListIndex)
         Case 11
            pnl_TraDep.Visible = True
            pnl_TraInd.Visible = False
            pnl_TraCom.Visible = False
            pnl_TraAcc.Visible = False
            pnl_TraRen.Visible = False
            
            Call fs_Activa_Dep(True)
            Call fs_Limpia_Dep
         
         Case 21
            pnl_TraDep.Visible = False
            pnl_TraInd.Visible = True
            pnl_TraCom.Visible = False
            pnl_TraAcc.Visible = False
            pnl_TraRen.Visible = False
            
            Call fs_Activa_Ind(True)
            Call fs_Limpia_Ind
            
         Case 31
            pnl_TraDep.Visible = False
            pnl_TraInd.Visible = False
            pnl_TraCom.Visible = True
            pnl_TraAcc.Visible = False
            pnl_TraRen.Visible = False
            
            Call fs_Activa_Com(True)
            Call fs_Limpia_Com
            
         Case 41
            pnl_TraDep.Visible = False
            pnl_TraInd.Visible = False
            pnl_TraCom.Visible = False
            pnl_TraAcc.Visible = True
            pnl_TraRen.Visible = False
            
            Call fs_Activa_Acc(True)
            Call fs_Limpia_Acc
            
         Case 51
            pnl_TraDep.Visible = False
            pnl_TraInd.Visible = False
            pnl_TraCom.Visible = False
            pnl_TraAcc.Visible = False
            pnl_TraRen.Visible = True
            
            Call fs_Activa_Ren(True)
            Call fs_Limpia_Ren
            
      End Select
      
      Call gs_SetFocus(cmb_OrdAct)
   End If
End Sub

Private Sub cmb_Com_RegTri_Click()
   Call gs_SetFocus(ipp_Com_PorPar)
End Sub

Private Sub cmb_Com_RegTri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Com_RegTri_Click
   End If
End Sub

Private Sub cmb_Com_TipLoc_Click()
   Call gs_SetFocus(cmd_Agrega)
   
   ipp_Com_AlqMen.Enabled = False
   txt_Com_NomArr.Enabled = False
   txt_Com_Tl1Arr.Enabled = False
   txt_Com_Tl2Arr.Enabled = False
   
   If cmb_Com_TipLoc.ListIndex > -1 Then
      If cmb_Com_TipLoc.ItemData(cmb_Com_TipLoc.ListIndex) = 2 Then
         ipp_Com_AlqMen.Enabled = True
         txt_Com_NomArr.Enabled = True
         txt_Com_Tl1Arr.Enabled = True
         txt_Com_Tl2Arr.Enabled = True
         
         Call gs_SetFocus(ipp_Com_AlqMen)
      Else
         txt_Com_NomArr.Text = ""
         txt_Com_Tl1Arr.Text = ""
         txt_Com_Tl2Arr.Text = ""
         ipp_Com_AlqMen.Value = 0
      End If
   End If
End Sub

Private Sub cmb_Com_TipLoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Com_TipLoc_Click
   End If
End Sub

Private Sub cmb_Dep_FreHab_Click()
   Call gs_SetFocus(cmb_Dep_NomCar)
End Sub

Private Sub cmb_Dep_FreHab_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Dep_FreHab_Click
   End If
End Sub

Private Sub cmb_Dep_TipDoc_Click()
   Call gs_SetFocus(txt_Dep_NumDoc)
End Sub

Private Sub cmb_Dep_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Dep_TipDoc_Click
   End If
End Sub

Private Sub cmb_Ind_ConLoc_Click()
   cmb_Ind_TDoEmp.Enabled = False
   txt_Ind_NDoEmp.Enabled = False
   txt_Ind_RazSoc.Enabled = False
   txt_Ind_Tl1Emp.Enabled = False
   txt_Ind_Tl2Emp.Enabled = False
   cmb_Ind_NomCar.Enabled = False
   txt_Ind_NomCar.Enabled = False
   ipp_Ind_FecIng.Enabled = False
   cmd_Ind_BusEmp.Enabled = False
   pnl_Ind_FlgEmp.Caption = ""
   
   Call gs_SetFocus(cmd_Agrega)
   
   If cmb_Ind_ConLoc.ListIndex > -1 Then
      If cmb_Ind_ConLoc.ItemData(cmb_Ind_ConLoc.ListIndex) = 1 Then
         cmb_Ind_TDoEmp.Enabled = True
         txt_Ind_NDoEmp.Enabled = True
         txt_Ind_RazSoc.Enabled = True
         txt_Ind_Tl1Emp.Enabled = True
         txt_Ind_Tl2Emp.Enabled = True
         cmb_Ind_NomCar.Enabled = True
         txt_Ind_NomCar.Enabled = True
         ipp_Ind_FecIng.Enabled = True
         cmd_Ind_BusEmp.Enabled = True
         
         Call gs_SetFocus(cmb_Ind_TDoEmp)
      Else
         cmb_Ind_TDoEmp.ListIndex = -1
         txt_Ind_NDoEmp.Text = ""
         txt_Ind_RazSoc.Text = ""
         txt_Ind_Tl1Emp.Text = ""
         txt_Ind_Tl2Emp.Text = ""
         cmb_Ind_NomCar.ListIndex = -1
         txt_Ind_NomCar.Text = ""
         ipp_Ind_FecIng.Text = Format(CDate(CDate(moddat_g_str_FecSis) - CDate(365)), "dd/mm/yyyy")
      End If
   End If
End Sub

Private Sub cmb_Ind_ConLoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Ind_ConLoc_Click
   End If
End Sub

Private Sub cmb_Ind_TDoEmp_Click()
   Call gs_SetFocus(txt_Ind_NDoEmp)
End Sub

Private Sub cmb_Ind_TDoEmp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Ind_TDoEmp_Click
   End If
End Sub

Private Sub cmb_Ind_TipDoc_Click()
   Call gs_SetFocus(txt_Ind_NumDoc)
End Sub

Private Sub cmb_Ind_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Ind_TipDoc_Click
   End If
End Sub

Private Sub cmb_OrdAct_Click()
   If cmb_ActEco.ListIndex > -1 Then
      If cmb_OrdAct.ListIndex > -1 Then
         If cmb_OrdAct.ItemData(cmb_OrdAct.ListIndex) = 9 And cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 11 Then
            ipp_Dep_FecCes.Enabled = True
         Else
            ipp_Dep_FecCes.Text = Format(CDate(moddat_g_str_FecSis) - CDate(365 * 2), "dd/mm/yyyy")
            ipp_Dep_FecCes.Enabled = False
         End If
      End If
      
      Select Case cmb_ActEco.ItemData(cmb_ActEco.ListIndex)
         Case 11:    Call gs_SetFocus(cmb_Dep_TipDoc)
         Case 21:    Call gs_SetFocus(cmb_Ind_TipDoc)
         Case 31:    Call gs_SetFocus(cmb_Com_TipDoc)
         Case 41:    Call gs_SetFocus(cmb_Acc_TipDoc)
         Case 51:    Call gs_SetFocus(cmb_Ren_TipDoc)
      End Select
   ElseIf cmb_ActEco.ListIndex = -1 And cmb_OrdAct.ListIndex > -1 Then
      Call gs_SetFocus(cmb_ActEco)
   End If
End Sub

Private Sub cmb_OrdAct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_OrdAct_Click
   End If
End Sub

Private Sub cmd_Acc_BusEmp_Click()
   If cmb_Acc_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Acc_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_Acc_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Acc_NumDoc)
      Exit Sub
   End If
   
   If Not gf_Valida_RUC(Mid(txt_Acc_NumDoc.Text, 1, Len(txt_Acc_NumDoc.Text) - 1), Right(txt_Acc_NumDoc.Text, 1)) Then
      MsgBox "El Número de Documento de Identidad no es válido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Acc_NumDoc)
      Exit Sub
   End If
   
   Call fs_BusEmp(cmb_Acc_TipDoc.ItemData(cmb_Acc_TipDoc.ListIndex), txt_Acc_NumDoc)
End Sub

Private Sub cmd_Agrega_Click()
   Dim r_int_Contad  As Integer
   
   If cmb_ActEco.ListIndex = -1 Then
      MsgBox "Seleccione la Actividad Económica.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ActEco)
      Exit Sub
   End If
   
   If cmb_OrdAct.ListIndex = -1 Then
      MsgBox "Seleccione el Orden de la Actividad Económica.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_OrdAct)
      Exit Sub
   End If

   Select Case cmb_ActEco.ItemData(cmb_ActEco.ListIndex)
      Case 11
         If Not ff_Valida_Dep() Then
            Exit Sub
         End If

      Case 21
         If Not ff_Valida_Ind() Then
            Exit Sub
         End If

      Case 31
         If Not ff_Valida_Com() Then
            Exit Sub
         End If
         
      Case 41
         If Not ff_Valida_Acc() Then
            Exit Sub
         End If

      Case 51
         If Not ff_Valida_Ren() Then
            Exit Sub
         End If
   End Select
   
   If MsgBox("¿Está seguro de agregar el Item a la Lista?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If l_int_FlgGrb = 1 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
   End If
   
   'Inicializando Columnas en Grid
   For r_int_Contad = 0 To grd_Listad.Cols - 1
      grd_Listad.Col = r_int_Contad
      grd_Listad.Text = ""
   Next r_int_Contad
   
   'Pasando Información de Campos a la Grid
   grd_Listad.Col = 0:  grd_Listad.Text = cmb_OrdAct.Text
   grd_Listad.Col = 1:  grd_Listad.Text = cmb_ActEco.Text
   grd_Listad.Col = 2:  grd_Listad.Text = cmb_OrdAct.ItemData(cmb_OrdAct.ListIndex)
   grd_Listad.Col = 3:  grd_Listad.Text = cmb_ActEco.ItemData(cmb_ActEco.ListIndex)

   'Generales
   Select Case cmb_ActEco.ItemData(cmb_ActEco.ListIndex)
      Case 11
         grd_Listad.Col = 4:  grd_Listad.Text = cmb_Dep_TipDoc.ItemData(cmb_Dep_TipDoc.ListIndex)
         grd_Listad.Col = 5:  grd_Listad.Text = txt_Dep_NumDoc.Text
         grd_Listad.Col = 6:  grd_Listad.Text = txt_Dep_RazSoc.Text
         grd_Listad.Col = 7:  grd_Listad.Text = txt_Dep_NomCom.Text
         grd_Listad.Col = 8:  grd_Listad.Text = l_arr_Dep_GirCom(cmb_Dep_GirCom.ListIndex + 1).Genera_TipPar
         grd_Listad.Col = 9:  grd_Listad.Text = l_arr_Dep_GirCom(cmb_Dep_GirCom.ListIndex + 1).Genera_Codigo
         grd_Listad.Col = 10: grd_Listad.Text = txt_Dep_GirCom.Text
         grd_Listad.Col = 11: grd_Listad.Text = txt_Dep_Sucurs.Text
         grd_Listad.Col = 12: grd_Listad.Text = cmb_Dep_TipVia.ItemData(cmb_Dep_TipVia.ListIndex)
         grd_Listad.Col = 13: grd_Listad.Text = txt_Dep_NomVia.Text
         grd_Listad.Col = 14: grd_Listad.Text = txt_Dep_Numero.Text
         grd_Listad.Col = 15: grd_Listad.Text = txt_Dep_Interi.Text
         grd_Listad.Col = 16: grd_Listad.Text = cmb_Dep_TipZon.ItemData(cmb_Dep_TipZon.ListIndex)
         grd_Listad.Col = 17: grd_Listad.Text = txt_Dep_NomZon.Text
         grd_Listad.Col = 18: grd_Listad.Text = cmb_Dep_DptDir.ItemData(cmb_Dep_DptDir.ListIndex)
         grd_Listad.Col = 19: grd_Listad.Text = cmb_Dep_PrvDir.ItemData(cmb_Dep_PrvDir.ListIndex)
         grd_Listad.Col = 20: grd_Listad.Text = cmb_Dep_DstDir.ItemData(cmb_Dep_DstDir.ListIndex)
         grd_Listad.Col = 21: grd_Listad.Text = txt_Dep_Refere.Text
         grd_Listad.Col = 22: grd_Listad.Text = txt_Dep_Telef1.Text
         grd_Listad.Col = 23: grd_Listad.Text = txt_Dep_Telef2.Text
         grd_Listad.Col = 24: grd_Listad.Text = txt_Dep_NumFax.Text
         grd_Listad.Col = 25: grd_Listad.Text = txt_Dep_TeleRH.Text
         grd_Listad.Col = 26: grd_Listad.Text = txt_Dep_AnexRH.Text
         grd_Listad.Col = 27: grd_Listad.Text = l_arr_Dep_GirCom(cmb_Dep_GirCom.ListIndex + 1).Genera_TipVal
         
         
         grd_Listad.Col = 29: grd_Listad.Text = pnl_Dep_FlgEmp.Caption
         grd_Listad.Col = 30: grd_Listad.Text = ipp_Dep_IngNet.Text
         grd_Listad.Col = 31: grd_Listad.Text = cmb_Dep_FreHab.ItemData(cmb_Dep_FreHab.ListIndex)
         grd_Listad.Col = 32: grd_Listad.Text = l_arr_Dep_NomCar(cmb_Dep_NomCar.ListIndex + 1).Genera_Codigo
         grd_Listad.Col = 33: grd_Listad.Text = txt_Dep_NomCar.Text
         grd_Listad.Col = 34: grd_Listad.Text = txt_Dep_NomAre.Text
         grd_Listad.Col = 35: grd_Listad.Text = ipp_Dep_FecIng.Text
         grd_Listad.Col = 36: grd_Listad.Text = txt_Dep_NumAnx.Text
         grd_Listad.Col = 37: grd_Listad.Text = txt_Dep_TelDir.Text
         grd_Listad.Col = 38: grd_Listad.Text = txt_Dep_Celula.Text
         grd_Listad.Col = 39: grd_Listad.Text = txt_Dep_DirEle.Text
         
         If ipp_Dep_FecCes.Enabled Then
            grd_Listad.Col = 40: grd_Listad.Text = ipp_Dep_FecCes.Text
         End If
         
      Case 21
         grd_Listad.Col = 4:  grd_Listad.Text = cmb_Ind_TipDoc.ItemData(cmb_Ind_TipDoc.ListIndex)
         grd_Listad.Col = 5:  grd_Listad.Text = txt_Ind_NumDoc.Text
         grd_Listad.Col = 8:  grd_Listad.Text = l_arr_Ind_GirCom(cmb_Ind_GirCom.ListIndex + 1).Genera_TipPar
         grd_Listad.Col = 9:  grd_Listad.Text = l_arr_Ind_GirCom(cmb_Ind_GirCom.ListIndex + 1).Genera_Codigo
         grd_Listad.Col = 10: grd_Listad.Text = txt_Ind_GirCom.Text
         grd_Listad.Col = 12: grd_Listad.Text = cmb_Ind_TipVia.ItemData(cmb_Ind_TipVia.ListIndex)
         grd_Listad.Col = 13: grd_Listad.Text = txt_Ind_NomVia.Text
         grd_Listad.Col = 14: grd_Listad.Text = txt_Ind_Numero.Text
         grd_Listad.Col = 15: grd_Listad.Text = txt_Ind_Interi.Text
         grd_Listad.Col = 16: grd_Listad.Text = cmb_Ind_TipZon.ItemData(cmb_Ind_TipZon.ListIndex)
         grd_Listad.Col = 17: grd_Listad.Text = txt_Ind_NomZon.Text
         grd_Listad.Col = 18: grd_Listad.Text = cmb_Ind_DptDir.ItemData(cmb_Ind_DptDir.ListIndex)
         grd_Listad.Col = 19: grd_Listad.Text = cmb_Ind_PrvDir.ItemData(cmb_Ind_PrvDir.ListIndex)
         grd_Listad.Col = 20: grd_Listad.Text = cmb_Ind_DstDir.ItemData(cmb_Ind_DstDir.ListIndex)
         grd_Listad.Col = 21: grd_Listad.Text = txt_Ind_Refere.Text
         grd_Listad.Col = 22: grd_Listad.Text = txt_Ind_Telef1.Text
         grd_Listad.Col = 23: grd_Listad.Text = txt_Ind_Telef2.Text
         grd_Listad.Col = 24: grd_Listad.Text = txt_Ind_NumFax.Text
         grd_Listad.Col = 27: grd_Listad.Text = l_arr_Ind_GirCom(cmb_Ind_GirCom.ListIndex + 1).Genera_TipVal
         
      
         grd_Listad.Col = 43: grd_Listad.Text = ipp_Ind_IngNet.Text
         grd_Listad.Col = 44: grd_Listad.Text = ipp_Ind_FecIni.Text
         grd_Listad.Col = 45: grd_Listad.Text = cmb_Ind_ConLoc.ItemData(cmb_Ind_ConLoc.ListIndex)
         
         If cmb_Ind_ConLoc.ItemData(cmb_Ind_ConLoc.ListIndex) = 1 Then
            grd_Listad.Col = 46: grd_Listad.Text = cmb_Ind_TDoEmp.ItemData(cmb_Ind_TDoEmp.ListIndex)
            grd_Listad.Col = 47: grd_Listad.Text = txt_Ind_NDoEmp.Text
            grd_Listad.Col = 48: grd_Listad.Text = txt_Ind_RazSoc.Text
            grd_Listad.Col = 49: grd_Listad.Text = txt_Ind_Tl1Emp.Text
            grd_Listad.Col = 50: grd_Listad.Text = txt_Ind_Tl2Emp.Text
            grd_Listad.Col = 51: grd_Listad.Text = l_arr_Ind_NomCar(cmb_Ind_NomCar.ListIndex + 1).Genera_Codigo
            grd_Listad.Col = 52: grd_Listad.Text = txt_Ind_NomCar.Text
            grd_Listad.Col = 53: grd_Listad.Text = ipp_Ind_FecIng.Text
            grd_Listad.Col = 54: grd_Listad.Text = pnl_Ind_FlgEmp.Caption
         End If
         
      Case 31
         grd_Listad.Col = 4:  grd_Listad.Text = cmb_Com_TipDoc.ItemData(cmb_Com_TipDoc.ListIndex)
         grd_Listad.Col = 5:  grd_Listad.Text = txt_Com_NumDoc.Text
         grd_Listad.Col = 6:  grd_Listad.Text = txt_Com_RazSoc.Text
         grd_Listad.Col = 7:  grd_Listad.Text = txt_Com_NomCom.Text
         grd_Listad.Col = 8:  grd_Listad.Text = l_arr_Com_GirCom(cmb_Com_GirCom.ListIndex + 1).Genera_TipPar
         grd_Listad.Col = 9:  grd_Listad.Text = l_arr_Com_GirCom(cmb_Com_GirCom.ListIndex + 1).Genera_Codigo
         grd_Listad.Col = 10: grd_Listad.Text = txt_Com_GirCom.Text
         grd_Listad.Col = 12: grd_Listad.Text = cmb_Com_TipVia.ItemData(cmb_Com_TipVia.ListIndex)
         grd_Listad.Col = 13: grd_Listad.Text = txt_Com_NomVia.Text
         grd_Listad.Col = 14: grd_Listad.Text = txt_Com_Numero.Text
         grd_Listad.Col = 15: grd_Listad.Text = txt_Com_Interi.Text
         grd_Listad.Col = 16: grd_Listad.Text = cmb_Com_TipZon.ItemData(cmb_Com_TipZon.ListIndex)
         grd_Listad.Col = 17: grd_Listad.Text = txt_Com_NomZon.Text
         grd_Listad.Col = 18: grd_Listad.Text = cmb_Com_DptDir.ItemData(cmb_Com_DptDir.ListIndex)
         grd_Listad.Col = 19: grd_Listad.Text = cmb_Com_PrvDir.ItemData(cmb_Com_PrvDir.ListIndex)
         grd_Listad.Col = 20: grd_Listad.Text = cmb_Com_DstDir.ItemData(cmb_Com_DstDir.ListIndex)
         grd_Listad.Col = 21: grd_Listad.Text = txt_Com_Refere.Text
         grd_Listad.Col = 22: grd_Listad.Text = txt_Com_Telef1.Text
         grd_Listad.Col = 23: grd_Listad.Text = txt_Com_Telef2.Text
         grd_Listad.Col = 24: grd_Listad.Text = txt_Com_NumFax.Text
         grd_Listad.Col = 27: grd_Listad.Text = l_arr_Com_GirCom(cmb_Com_GirCom.ListIndex + 1).Genera_TipVal
      
         grd_Listad.Col = 85: grd_Listad.Text = ipp_Com_IngNet.Text
         grd_Listad.Col = 86: grd_Listad.Text = ipp_Com_VtaMen.Text
         grd_Listad.Col = 87: grd_Listad.Text = ipp_Com_FecIni.Text
         grd_Listad.Col = 88: grd_Listad.Text = cmb_Com_RegTri.ItemData(cmb_Com_RegTri.ListIndex)
         grd_Listad.Col = 89: grd_Listad.Text = ipp_Com_PorPar.Text
         grd_Listad.Col = 90: grd_Listad.Text = cmb_Com_TipLoc.ItemData(cmb_Com_TipLoc.ListIndex)
         
         If cmb_Com_TipLoc.ItemData(cmb_Com_TipLoc.ListIndex) = 2 Then
            grd_Listad.Col = 91: grd_Listad.Text = ipp_Com_AlqMen.Text
            grd_Listad.Col = 92: grd_Listad.Text = txt_Com_NomArr.Text
            grd_Listad.Col = 93: grd_Listad.Text = txt_Com_Tl1Arr.Text
            grd_Listad.Col = 94: grd_Listad.Text = txt_Com_Tl2Arr.Text
         End If
         
         grd_Listad.Col = 95: grd_Listad.Text = pnl_Com_FlgEmp.Caption
         
      Case 41
         grd_Listad.Col = 4:  grd_Listad.Text = cmb_Acc_TipDoc.ItemData(cmb_Acc_TipDoc.ListIndex)
         grd_Listad.Col = 5:  grd_Listad.Text = txt_Acc_NumDoc.Text
         grd_Listad.Col = 6:  grd_Listad.Text = txt_Acc_RazSoc.Text
         grd_Listad.Col = 7:  grd_Listad.Text = txt_Acc_NomCom.Text
         grd_Listad.Col = 8:  grd_Listad.Text = l_arr_Acc_GirCom(cmb_Acc_GirCom.ListIndex + 1).Genera_TipPar
         grd_Listad.Col = 9:  grd_Listad.Text = l_arr_Acc_GirCom(cmb_Acc_GirCom.ListIndex + 1).Genera_Codigo
         grd_Listad.Col = 10: grd_Listad.Text = txt_Acc_GirCom.Text
         grd_Listad.Col = 12: grd_Listad.Text = cmb_Acc_TipVia.ItemData(cmb_Acc_TipVia.ListIndex)
         grd_Listad.Col = 13: grd_Listad.Text = txt_Acc_NomVia.Text
         grd_Listad.Col = 14: grd_Listad.Text = txt_Acc_Numero.Text
         grd_Listad.Col = 15: grd_Listad.Text = txt_Acc_Interi.Text
         grd_Listad.Col = 16: grd_Listad.Text = cmb_Acc_TipZon.ItemData(cmb_Acc_TipZon.ListIndex)
         grd_Listad.Col = 17: grd_Listad.Text = txt_Acc_NomZon.Text
         grd_Listad.Col = 18: grd_Listad.Text = cmb_Acc_DptDir.ItemData(cmb_Acc_DptDir.ListIndex)
         grd_Listad.Col = 19: grd_Listad.Text = cmb_Acc_PrvDir.ItemData(cmb_Acc_PrvDir.ListIndex)
         grd_Listad.Col = 20: grd_Listad.Text = cmb_Acc_DstDir.ItemData(cmb_Acc_DstDir.ListIndex)
         grd_Listad.Col = 21: grd_Listad.Text = txt_Acc_Refere.Text
         grd_Listad.Col = 22: grd_Listad.Text = txt_Acc_Telef1.Text
         grd_Listad.Col = 23: grd_Listad.Text = txt_Acc_Telef2.Text
         grd_Listad.Col = 24: grd_Listad.Text = txt_Acc_NumFax.Text
         grd_Listad.Col = 27: grd_Listad.Text = l_arr_Acc_GirCom(cmb_Acc_GirCom.ListIndex + 1).Genera_TipVal
      
         grd_Listad.Col = 57: grd_Listad.Text = ipp_Acc_IngNet.Text
         grd_Listad.Col = 58: grd_Listad.Text = ipp_Acc_PorAcc.Text
         grd_Listad.Col = 59: grd_Listad.Text = ipp_Acc_FecAnt.Text
         
         grd_Listad.Col = 60: grd_Listad.Text = pnl_Acc_FlgEmp.Caption
      
      Case 51
         grd_Listad.Col = 4:  grd_Listad.Text = cmb_Ren_TipDoc.ItemData(cmb_Ren_TipDoc.ListIndex)
         grd_Listad.Col = 5:  grd_Listad.Text = txt_Ren_NumDoc.Text
         grd_Listad.Col = 8:  grd_Listad.Text = l_arr_Ren_GirCom(cmb_Ren_GirCom.ListIndex + 1).Genera_TipPar
         grd_Listad.Col = 9:  grd_Listad.Text = l_arr_Ren_GirCom(cmb_Ren_GirCom.ListIndex + 1).Genera_Codigo
         grd_Listad.Col = 27: grd_Listad.Text = l_arr_Ren_GirCom(cmb_Ren_GirCom.ListIndex + 1).Genera_TipVal
      
         grd_Listad.Col = 62: grd_Listad.Text = txt_Ren_Direc1.Text
         grd_Listad.Col = 63: grd_Listad.Text = txt_Ren_NomAr1.Text
         grd_Listad.Col = 64: grd_Listad.Text = txt_Ren_Tele11.Text
         grd_Listad.Col = 65: grd_Listad.Text = txt_Ren_Tele21.Text
         grd_Listad.Col = 66: grd_Listad.Text = ipp_Ren_AlqMe1.Text
         grd_Listad.Col = 67: grd_Listad.Text = ipp_Ren_FIAlq1.Text
         
         grd_Listad.Col = 68: grd_Listad.Text = chk_Alqui2.Value
         If chk_Alqui2.Value = 1 Then
            grd_Listad.Col = 69: grd_Listad.Text = txt_Ren_Direc2.Text
            grd_Listad.Col = 70: grd_Listad.Text = txt_Ren_NomAr2.Text
            grd_Listad.Col = 71: grd_Listad.Text = txt_Ren_Tele12.Text
            grd_Listad.Col = 72: grd_Listad.Text = txt_Ren_Tele22.Text
            grd_Listad.Col = 73: grd_Listad.Text = ipp_Ren_AlqMe2.Text
            grd_Listad.Col = 74: grd_Listad.Text = ipp_Ren_FIAlq2.Text
         End If
         
         grd_Listad.Col = 75: grd_Listad.Text = chk_Alqui3.Value
         If chk_Alqui3.Value = 1 Then
            grd_Listad.Col = 76: grd_Listad.Text = txt_Ren_Direc3.Text
            grd_Listad.Col = 77: grd_Listad.Text = txt_Ren_NomAr3.Text
            grd_Listad.Col = 78: grd_Listad.Text = txt_Ren_Tele13.Text
            grd_Listad.Col = 79: grd_Listad.Text = txt_Ren_Tele23.Text
            grd_Listad.Col = 80: grd_Listad.Text = ipp_Ren_AlqMe3.Text
            grd_Listad.Col = 81: grd_Listad.Text = ipp_Ren_FIAlq3.Text
         End If
         
         grd_Listad.Col = 82: grd_Listad.Text = ipp_Ren_IngNet.Text
   End Select
   
   Call gs_UbiIniGrid(grd_Listad)
   
   Call fs_Activa(False)
   Call fs_Limpia
   
   Call fs_Activa_Dep(False)
   Call fs_Limpia_Dep
   
   pnl_TraDep.Visible = True
   
   Call gs_RefrescaGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
   
   cmd_Grabar.Enabled = True
   cmd_EdiAct.Enabled = True
   cmd_BorAct.Enabled = True
End Sub

Private Sub cmd_BorAct_Click()
   If MsgBox("¿Está seguro de eliminar la actividad?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If grd_Listad.Rows = 1 Then
      grd_Listad.Rows = 0
   Else
      grd_Listad.RemoveItem grd_Listad.Row
   End If
   
   If grd_Listad.Rows = 0 Then
      cmd_BorAct.Enabled = False
      cmd_EdiAct.Enabled = False
   End If
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Activa(False)
   Call fs_Limpia
   
   pnl_TraDep.Visible = True
   Call fs_Activa_Dep(False)
   
   If grd_Listad.Rows > 0 Then
      Call gs_RefrescaGrid(grd_Listad)
      Call gs_SetFocus(grd_Listad)
   
      cmd_Grabar.Enabled = True
      cmd_EdiAct.Enabled = True
      cmd_BorAct.Enabled = True
   Else
      Call gs_SetFocus(cmd_NueAct)
   End If
End Sub

Private Sub cmd_Com_BusEmp_Click()
   If cmb_Com_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Com_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_Com_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Com_NumDoc)
      Exit Sub
   End If
   
   If Not gf_Valida_RUC(Mid(txt_Com_NumDoc.Text, 1, Len(txt_Com_NumDoc.Text) - 1), Right(txt_Com_NumDoc.Text, 1)) Then
      MsgBox "El Número de Documento de Identidad no es válido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Com_NumDoc)
      Exit Sub
   End If
   
   Call fs_BusEmp(cmb_Com_TipDoc.ItemData(cmb_Com_TipDoc.ListIndex), txt_Com_NumDoc)
End Sub

Private Sub cmd_Dep_BusEmp_Click()
   If cmb_Dep_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Dep_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_Dep_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Dep_NumDoc)
      Exit Sub
   End If
   
   If Not gf_Valida_RUC(Mid(txt_Dep_NumDoc.Text, 1, Len(txt_Dep_NumDoc.Text) - 1), Right(txt_Dep_NumDoc.Text, 1)) Then
      MsgBox "El Número de Documento de Identidad no es válido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Dep_NumDoc)
      Exit Sub
   End If
   
   Call fs_BusEmp(cmb_Dep_TipDoc.ItemData(cmb_Dep_TipDoc.ListIndex), txt_Dep_NumDoc)
End Sub

Private Sub cmd_EdiAct_Click()
   Dim r_int_ActEco     As Integer
   
   l_int_FlgGrb = 2
   
   grd_Listad.Col = 3
   r_int_ActEco = CInt(grd_Listad.Text)
   
   Call fs_Activa(True)
   Call fs_Limpia
   
   Call gs_BuscarCombo_Item(cmb_ActEco, r_int_ActEco)
   
   grd_Listad.Col = 2
   Call gs_BuscarCombo_Item(cmb_OrdAct, CInt(grd_Listad.Text))
   
   Select Case r_int_ActEco
      Case 11
         Call fs_Activa_Dep(True)
         Call fs_Limpia_Dep
         
         grd_Listad.Col = 4
         Call gs_BuscarCombo_Item(cmb_Dep_TipDoc, CInt(grd_Listad.Text))
         
         grd_Listad.Col = 5
         txt_Dep_NumDoc.Text = grd_Listad.Text
         
         'Obteniendo Información Actualizada de Empresas
         Call fs_BusEmp(cmb_Dep_TipDoc.ItemData(cmb_Dep_TipDoc.ListIndex), txt_Dep_NumDoc)
         
         If pnl_Dep_FlgEmp = "NR" Then
            'Si la Empresa no está registrada
            
            grd_Listad.Col = 6
            txt_Dep_RazSoc.Text = grd_Listad.Text
            
            grd_Listad.Col = 7
            txt_Dep_NomCom.Text = grd_Listad.Text
         
            grd_Listad.Col = 9
            cmb_Dep_GirCom.ListIndex = gf_Busca_Arregl(l_arr_Dep_GirCom, grd_Listad.Text) - 1
         
            grd_Listad.Col = 10
            txt_Dep_GirCom.Text = grd_Listad.Text
         
            If l_arr_Dep_GirCom(cmb_Dep_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
               txt_Dep_GirCom.Enabled = True
            End If
         
            grd_Listad.Col = 11
            txt_Dep_Sucurs.Text = grd_Listad.Text
         
            If Len(Trim(txt_Dep_Sucurs.Text)) > 0 Then
               chk_Dep_Sucurs.Value = 1
            End If
         
            grd_Listad.Col = 12
            Call gs_BuscarCombo_Item(cmb_Dep_TipVia, CInt(grd_Listad.Text))
            
            grd_Listad.Col = 13
            txt_Dep_NomVia.Text = grd_Listad.Text
            
            grd_Listad.Col = 14
            txt_Dep_Numero.Text = grd_Listad.Text
            
            grd_Listad.Col = 15
            txt_Dep_Interi.Text = grd_Listad.Text
         
            grd_Listad.Col = 16
            Call gs_BuscarCombo_Item(cmb_Dep_TipZon, CInt(grd_Listad.Text))
         
            grd_Listad.Col = 17
            txt_Dep_NomZon.Text = grd_Listad.Text
         
            grd_Listad.Col = 18
            Call gs_BuscarCombo_Item(cmb_Dep_DptDir, CInt(grd_Listad.Text))
            
            Call moddat_gs_Carga_Provin(cmb_Dep_PrvDir, Format(cmb_Dep_DptDir.ItemData(cmb_Dep_DptDir.ListIndex), "00"))
            
            grd_Listad.Col = 19
            Call gs_BuscarCombo_Item(cmb_Dep_PrvDir, CInt(grd_Listad.Text))
            
            Call moddat_gs_Carga_Distri(cmb_Dep_DstDir, Format(cmb_Dep_DptDir.ItemData(cmb_Dep_DptDir.ListIndex), "00"), Format(cmb_Dep_PrvDir.ItemData(cmb_Dep_PrvDir.ListIndex), "00"))
            
            grd_Listad.Col = 20
            Call gs_BuscarCombo_Item(cmb_Dep_DstDir, CInt(grd_Listad.Text))
            
            grd_Listad.Col = 21
            txt_Dep_Refere.Text = grd_Listad.Text
            
            grd_Listad.Col = 22
            txt_Dep_Telef1.Text = grd_Listad.Text
            
            grd_Listad.Col = 23
            txt_Dep_Telef2.Text = grd_Listad.Text
            
            grd_Listad.Col = 24
            txt_Dep_NumFax.Text = grd_Listad.Text
         
            grd_Listad.Col = 25
            txt_Dep_TeleRH.Text = grd_Listad.Text
         
            grd_Listad.Col = 26
            txt_Dep_AnexRH.Text = grd_Listad.Text
         
            grd_Listad.Col = 29
            pnl_Dep_FlgEmp.Caption = grd_Listad.Text
         Else
            'Si la Empresa está registrada sólo obtiene información siempre y cuando sea Trabajador de Sucursal
            
            grd_Listad.Col = 11
            txt_Dep_Sucurs.Text = grd_Listad.Text
         
            If Len(Trim(txt_Dep_Sucurs.Text)) > 0 Then
               chk_Dep_Sucurs.Value = 1
         
               grd_Listad.Col = 12
               Call gs_BuscarCombo_Item(cmb_Dep_TipVia, CInt(grd_Listad.Text))
               
               grd_Listad.Col = 13
               txt_Dep_NomVia.Text = grd_Listad.Text
               
               grd_Listad.Col = 14
               txt_Dep_Numero.Text = grd_Listad.Text
               
               grd_Listad.Col = 15
               txt_Dep_Interi.Text = grd_Listad.Text
            
               grd_Listad.Col = 16
               Call gs_BuscarCombo_Item(cmb_Dep_TipZon, CInt(grd_Listad.Text))
            
               grd_Listad.Col = 17
               txt_Dep_NomZon.Text = grd_Listad.Text
            
               grd_Listad.Col = 18
               Call gs_BuscarCombo_Item(cmb_Dep_DptDir, CInt(grd_Listad.Text))
               
               Call moddat_gs_Carga_Provin(cmb_Dep_PrvDir, Format(cmb_Dep_DptDir.ItemData(cmb_Dep_DptDir.ListIndex), "00"))
               
               grd_Listad.Col = 19
               Call gs_BuscarCombo_Item(cmb_Dep_PrvDir, CInt(grd_Listad.Text))
               
               Call moddat_gs_Carga_Distri(cmb_Dep_DstDir, Format(cmb_Dep_DptDir.ItemData(cmb_Dep_DptDir.ListIndex), "00"), Format(cmb_Dep_PrvDir.ItemData(cmb_Dep_PrvDir.ListIndex), "00"))
               
               grd_Listad.Col = 20
               Call gs_BuscarCombo_Item(cmb_Dep_DstDir, CInt(grd_Listad.Text))
               
               grd_Listad.Col = 21
               txt_Dep_Refere.Text = grd_Listad.Text
               
               grd_Listad.Col = 22
               txt_Dep_Telef1.Text = grd_Listad.Text
               
               grd_Listad.Col = 23
               txt_Dep_Telef2.Text = grd_Listad.Text
               
               grd_Listad.Col = 24
               txt_Dep_NumFax.Text = grd_Listad.Text
            End If
         End If
         
         grd_Listad.Col = 30
         ipp_Dep_IngNet.Value = CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 31
         Call gs_BuscarCombo_Item(cmb_Dep_FreHab, CInt(grd_Listad.Text))
         
         grd_Listad.Col = 32
         cmb_Dep_NomCar.ListIndex = gf_Busca_Arregl(l_arr_Dep_NomCar, grd_Listad.Text) - 1
         
         grd_Listad.Col = 33
         txt_Dep_NomCar.Text = grd_Listad.Text
         
         grd_Listad.Col = 34
         txt_Dep_NomAre.Text = grd_Listad.Text
         
         grd_Listad.Col = 35
         ipp_Dep_FecIng.Text = grd_Listad.Text
         
         grd_Listad.Col = 36
         txt_Dep_NumAnx.Text = grd_Listad.Text
         
         grd_Listad.Col = 37
         txt_Dep_TelDir.Text = grd_Listad.Text
         
         grd_Listad.Col = 38
         txt_Dep_Celula.Text = grd_Listad.Text
         
         grd_Listad.Col = 39
         txt_Dep_DirEle.Text = grd_Listad.Text
         
         If cmb_OrdAct.ItemData(cmb_OrdAct.ListIndex) = 9 Then
            grd_Listad.Col = 40
            ipp_Dep_FecCes.Text = grd_Listad.Text
         End If
         
         If pnl_Dep_FlgEmp.Caption <> "NR" Then
            txt_Dep_RazSoc.Enabled = False
            txt_Dep_NomCom.Enabled = False
            cmb_Dep_GirCom.Enabled = False
            txt_Dep_GirCom.Enabled = False
            
            If chk_Dep_Sucurs.Value <> 1 Then
               cmb_Dep_TipVia.Enabled = False
               txt_Dep_NomVia.Enabled = False
               txt_Dep_Numero.Enabled = False
               txt_Dep_Interi.Enabled = False
               cmb_Dep_TipZon.Enabled = False
               txt_Dep_NomZon.Enabled = False
               cmb_Dep_DptDir.Enabled = False
               cmb_Dep_PrvDir.Enabled = False
               cmb_Dep_DstDir.Enabled = False
               txt_Dep_Refere.Enabled = False
               txt_Dep_Telef1.Enabled = False
               txt_Dep_Telef2.Enabled = False
               txt_Dep_NumFax.Enabled = False
               txt_Dep_TeleRH.Enabled = False
               txt_Dep_AnexRH.Enabled = False
            End If
         End If
         
      Case 21
         Call fs_Activa_Ind(True)
         Call fs_Limpia_Ind
         
         grd_Listad.Col = 4
         Call gs_BuscarCombo_Item(cmb_Ind_TipDoc, CInt(grd_Listad.Text))
         
         grd_Listad.Col = 5
         txt_Ind_NumDoc.Text = grd_Listad.Text
         
         grd_Listad.Col = 9
         cmb_Ind_GirCom.ListIndex = gf_Busca_Arregl(l_arr_Ind_GirCom, grd_Listad.Text) - 1
         
         grd_Listad.Col = 10
         txt_Ind_GirCom.Text = grd_Listad.Text
         
         If l_arr_Ind_GirCom(cmb_Ind_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
            txt_Ind_GirCom.Enabled = True
         End If
         
         grd_Listad.Col = 12
         Call gs_BuscarCombo_Item(cmb_Ind_TipVia, CInt(grd_Listad.Text))
         
         grd_Listad.Col = 13
         txt_Ind_NomVia.Text = grd_Listad.Text
         
         grd_Listad.Col = 14
         txt_Ind_Numero.Text = grd_Listad.Text
         
         grd_Listad.Col = 15
         txt_Ind_Interi.Text = grd_Listad.Text
      
         grd_Listad.Col = 16
         Call gs_BuscarCombo_Item(cmb_Ind_TipZon, CInt(grd_Listad.Text))
      
         grd_Listad.Col = 17
         txt_Ind_NomZon.Text = grd_Listad.Text
      
         grd_Listad.Col = 18
         Call gs_BuscarCombo_Item(cmb_Ind_DptDir, CInt(grd_Listad.Text))
         
         Call moddat_gs_Carga_Provin(cmb_Ind_PrvDir, Format(cmb_Ind_DptDir.ItemData(cmb_Ind_DptDir.ListIndex), "00"))
         
         grd_Listad.Col = 19
         Call gs_BuscarCombo_Item(cmb_Ind_PrvDir, CInt(grd_Listad.Text))
         
         Call moddat_gs_Carga_Distri(cmb_Ind_DstDir, Format(cmb_Ind_DptDir.ItemData(cmb_Ind_DptDir.ListIndex), "00"), Format(cmb_Ind_PrvDir.ItemData(cmb_Ind_PrvDir.ListIndex), "00"))
         
         grd_Listad.Col = 20
         Call gs_BuscarCombo_Item(cmb_Ind_DstDir, CInt(grd_Listad.Text))
         
         grd_Listad.Col = 21
         txt_Ind_Refere.Text = grd_Listad.Text
         
         grd_Listad.Col = 22
         txt_Ind_Telef1.Text = grd_Listad.Text
         
         grd_Listad.Col = 23
         txt_Ind_Telef2.Text = grd_Listad.Text
         
         grd_Listad.Col = 24
         txt_Ind_NumFax.Text = grd_Listad.Text
         
         grd_Listad.Col = 43
         ipp_Ind_IngNet.Value = CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 44
         ipp_Ind_FecIni.Text = grd_Listad.Text
         
         grd_Listad.Col = 45
         Call gs_BuscarCombo_Item(cmb_Ind_ConLoc, CInt(grd_Listad.Text))
         
         If cmb_Ind_ConLoc.ItemData(cmb_Ind_ConLoc.ListIndex) = 1 Then
            grd_Listad.Col = 46
            Call gs_BuscarCombo_Item(cmb_Ind_TDoEmp, CInt(grd_Listad.Text))
            
            grd_Listad.Col = 47
            txt_Ind_NDoEmp.Text = grd_Listad.Text
            
            'Obteniendo Información Actualizada de Empresas
            Call fs_BusEmp(cmb_Ind_TDoEmp.ItemData(cmb_Ind_TDoEmp.ListIndex), txt_Ind_NDoEmp)
            
            If pnl_Ind_FlgEmp = "NR" Then
               grd_Listad.Col = 48
               txt_Ind_RazSoc.Text = grd_Listad.Text
               
               grd_Listad.Col = 49
               txt_Ind_Tl1Emp.Text = grd_Listad.Text
               
               grd_Listad.Col = 50
               txt_Ind_Tl2Emp.Text = grd_Listad.Text
            End If
            
            grd_Listad.Col = 51
            cmb_Ind_NomCar.ListIndex = gf_Busca_Arregl(l_arr_Ind_NomCar, grd_Listad.Text) - 1
            
            grd_Listad.Col = 52
            txt_Ind_NomCar.Text = grd_Listad.Text
         
            grd_Listad.Col = 53
            ipp_Ind_FecIng.Text = grd_Listad.Text
         
            grd_Listad.Col = 54
            pnl_Ind_FlgEmp.Caption = grd_Listad.Text
         
            If pnl_Ind_FlgEmp.Caption <> "NR" Then
               txt_Ind_RazSoc.Enabled = False
               txt_Ind_Tl1Emp.Enabled = False
               txt_Ind_Tl2Emp.Enabled = False
            End If
         End If
         
         
      Case 31
         Call fs_Activa_Com(True)
         Call fs_Limpia_Com
         
         grd_Listad.Col = 4
         Call gs_BuscarCombo_Item(cmb_Com_TipDoc, CInt(grd_Listad.Text))
         
         grd_Listad.Col = 5
         txt_Com_NumDoc.Text = grd_Listad.Text
         
         'Obteniendo Información Actualizada de Empresas
         Call fs_BusEmp(cmb_Com_TipDoc.ItemData(cmb_Com_TipDoc.ListIndex), txt_Com_NumDoc)
         
         If pnl_Com_FlgEmp = "NR" Then
            grd_Listad.Col = 6
            txt_Com_RazSoc.Text = grd_Listad.Text
            
            grd_Listad.Col = 7
            txt_Com_NomCom.Text = grd_Listad.Text
            
            grd_Listad.Col = 9
            cmb_Com_GirCom.ListIndex = gf_Busca_Arregl(l_arr_Com_GirCom, grd_Listad.Text) - 1
            
            grd_Listad.Col = 10
            txt_Com_GirCom.Text = grd_Listad.Text
            
            If l_arr_Com_GirCom(cmb_Com_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
               txt_Com_GirCom.Enabled = True
            End If
            
            grd_Listad.Col = 12
            Call gs_BuscarCombo_Item(cmb_Com_TipVia, CInt(grd_Listad.Text))
            
            grd_Listad.Col = 13
            txt_Com_NomVia.Text = grd_Listad.Text
            
            grd_Listad.Col = 14
            txt_Com_Numero.Text = grd_Listad.Text
            
            grd_Listad.Col = 15
            txt_Com_Interi.Text = grd_Listad.Text
         
            grd_Listad.Col = 16
            Call gs_BuscarCombo_Item(cmb_Com_TipZon, CInt(grd_Listad.Text))
         
            grd_Listad.Col = 17
            txt_Com_NomZon.Text = grd_Listad.Text
         
            grd_Listad.Col = 18
            Call gs_BuscarCombo_Item(cmb_Com_DptDir, CInt(grd_Listad.Text))
            
            Call moddat_gs_Carga_Provin(cmb_Com_PrvDir, Format(cmb_Com_DptDir.ItemData(cmb_Com_DptDir.ListIndex), "00"))
            
            grd_Listad.Col = 19
            Call gs_BuscarCombo_Item(cmb_Com_PrvDir, CInt(grd_Listad.Text))
            
            Call moddat_gs_Carga_Distri(cmb_Com_DstDir, Format(cmb_Com_DptDir.ItemData(cmb_Com_DptDir.ListIndex), "00"), Format(cmb_Com_PrvDir.ItemData(cmb_Com_PrvDir.ListIndex), "00"))
            
            grd_Listad.Col = 20
            Call gs_BuscarCombo_Item(cmb_Com_DstDir, CInt(grd_Listad.Text))
            
            grd_Listad.Col = 21
            txt_Com_Refere.Text = grd_Listad.Text
            
            grd_Listad.Col = 22
            txt_Com_Telef1.Text = grd_Listad.Text
            
            grd_Listad.Col = 23
            txt_Com_Telef2.Text = grd_Listad.Text
            
            grd_Listad.Col = 24
            txt_Com_NumFax.Text = grd_Listad.Text
         End If
         
         grd_Listad.Col = 85
         ipp_Com_IngNet.Value = CDbl(grd_Listad.Text)
      
         grd_Listad.Col = 86
         ipp_Com_VtaMen.Value = CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 87
         ipp_Com_FecIni.Text = grd_Listad.Text
         
         grd_Listad.Col = 88
         Call gs_BuscarCombo_Item(cmb_Com_RegTri, CInt(grd_Listad.Text))
         
         grd_Listad.Col = 89
         ipp_Com_PorPar.Value = CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 90
         Call gs_BuscarCombo_Item(cmb_Com_TipLoc, CInt(grd_Listad.Text))
         
         If cmb_Com_TipLoc.ItemData(cmb_Com_TipLoc.ListIndex) = 2 Then
            grd_Listad.Col = 91
            ipp_Com_AlqMen.Value = CDbl(grd_Listad.Text)
         
            grd_Listad.Col = 92
            txt_Com_NomArr.Text = grd_Listad.Text
         
            grd_Listad.Col = 93
            txt_Com_Tl1Arr.Text = grd_Listad.Text
         
            grd_Listad.Col = 94
            txt_Com_Tl2Arr.Text = grd_Listad.Text
         End If
         
         grd_Listad.Col = 95
         pnl_Com_FlgEmp.Caption = grd_Listad.Text
      
         If pnl_Com_FlgEmp.Caption <> "NR" Then
            txt_Com_RazSoc.Enabled = False
            txt_Com_NomCom.Enabled = False
            cmb_Com_GirCom.Enabled = False
            txt_Com_GirCom.Enabled = False
            
            cmb_Com_TipVia.Enabled = False
            txt_Com_NomVia.Enabled = False
            txt_Com_Numero.Enabled = False
            txt_Com_Interi.Enabled = False
            cmb_Com_TipZon.Enabled = False
            txt_Com_NomZon.Enabled = False
            cmb_Com_DptDir.Enabled = False
            cmb_Com_PrvDir.Enabled = False
            cmb_Com_DstDir.Enabled = False
            txt_Com_Refere.Enabled = False
            txt_Com_Telef1.Enabled = False
            txt_Com_Telef2.Enabled = False
            txt_Com_NumFax.Enabled = False
         End If
      
      Case 41
         Call fs_Activa_Acc(True)
         Call fs_Limpia_Acc
         
         grd_Listad.Col = 4
         Call gs_BuscarCombo_Item(cmb_Acc_TipDoc, CInt(grd_Listad.Text))
         
         grd_Listad.Col = 5
         txt_Acc_NumDoc.Text = grd_Listad.Text
         
         'Obteniendo Información Actualizada de Empresas
         Call fs_BusEmp(cmb_Acc_TipDoc.ItemData(cmb_Acc_TipDoc.ListIndex), txt_Acc_NumDoc)
         
         If pnl_Acc_FlgEmp = "NR" Then
            grd_Listad.Col = 6
            txt_Acc_RazSoc.Text = grd_Listad.Text
            
            grd_Listad.Col = 7
            txt_Acc_NomCom.Text = grd_Listad.Text
           
            grd_Listad.Col = 9
            cmb_Acc_GirCom.ListIndex = gf_Busca_Arregl(l_arr_Acc_GirCom, grd_Listad.Text) - 1
            
            grd_Listad.Col = 10
            txt_Acc_GirCom.Text = grd_Listad.Text
            
            If l_arr_Acc_GirCom(cmb_Acc_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
               txt_Acc_GirCom.Enabled = True
            End If
            
            grd_Listad.Col = 12
            Call gs_BuscarCombo_Item(cmb_Acc_TipVia, CInt(grd_Listad.Text))
            
            grd_Listad.Col = 13
            txt_Acc_NomVia.Text = grd_Listad.Text
            
            grd_Listad.Col = 14
            txt_Acc_Numero.Text = grd_Listad.Text
            
            grd_Listad.Col = 15
            txt_Acc_Interi.Text = grd_Listad.Text
         
            grd_Listad.Col = 16
            Call gs_BuscarCombo_Item(cmb_Acc_TipZon, CInt(grd_Listad.Text))
         
            grd_Listad.Col = 17
            txt_Acc_NomZon.Text = grd_Listad.Text
         
            grd_Listad.Col = 18
            Call gs_BuscarCombo_Item(cmb_Acc_DptDir, CInt(grd_Listad.Text))
            
            Call moddat_gs_Carga_Provin(cmb_Acc_PrvDir, Format(cmb_Acc_DptDir.ItemData(cmb_Acc_DptDir.ListIndex), "00"))
            
            grd_Listad.Col = 19
            Call gs_BuscarCombo_Item(cmb_Acc_PrvDir, CInt(grd_Listad.Text))
            
            Call moddat_gs_Carga_Distri(cmb_Acc_DstDir, Format(cmb_Acc_DptDir.ItemData(cmb_Acc_DptDir.ListIndex), "00"), Format(cmb_Acc_PrvDir.ItemData(cmb_Acc_PrvDir.ListIndex), "00"))
            
            grd_Listad.Col = 20
            Call gs_BuscarCombo_Item(cmb_Acc_DstDir, CInt(grd_Listad.Text))
            
            grd_Listad.Col = 21
            txt_Acc_Refere.Text = grd_Listad.Text
            
            grd_Listad.Col = 22
            txt_Acc_Telef1.Text = grd_Listad.Text
            
            grd_Listad.Col = 23
            txt_Acc_Telef2.Text = grd_Listad.Text
            
            grd_Listad.Col = 24
            txt_Acc_NumFax.Text = grd_Listad.Text
         End If
         
         grd_Listad.Col = 57
         ipp_Acc_IngNet.Value = CDbl(grd_Listad.Text)
      
         grd_Listad.Col = 58
         ipp_Acc_PorAcc.Value = CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 59
         ipp_Acc_FecAnt.Text = grd_Listad.Text
         
         grd_Listad.Col = 60
         pnl_Acc_FlgEmp.Caption = grd_Listad.Text
         
         If pnl_Acc_FlgEmp.Caption <> "NR" Then
            txt_Acc_RazSoc.Enabled = False
            txt_Acc_NomCom.Enabled = False
            cmb_Acc_GirCom.Enabled = False
            txt_Acc_GirCom.Enabled = False

            cmb_Acc_TipVia.Enabled = False
            txt_Acc_NomVia.Enabled = False
            txt_Acc_Numero.Enabled = False
            txt_Acc_Interi.Enabled = False
            cmb_Acc_TipZon.Enabled = False
            txt_Acc_NomZon.Enabled = False
            cmb_Acc_DptDir.Enabled = False
            cmb_Acc_PrvDir.Enabled = False
            cmb_Acc_DstDir.Enabled = False
            txt_Acc_Refere.Enabled = False
            txt_Acc_Telef1.Enabled = False
            txt_Acc_Telef2.Enabled = False
            txt_Acc_NumFax.Enabled = False
         End If
         
      Case 51
         Call fs_Activa_Ren(True)
         Call fs_Limpia_Ren
         
         grd_Listad.Col = 4
         Call gs_BuscarCombo_Item(cmb_Ren_TipDoc, CInt(grd_Listad.Text))
         
         grd_Listad.Col = 5
         txt_Ren_NumDoc.Text = grd_Listad.Text
         
         grd_Listad.Col = 9
         cmb_Ren_GirCom.ListIndex = gf_Busca_Arregl(l_arr_Ren_GirCom, grd_Listad.Text) - 1
         
         grd_Listad.Col = 62
         txt_Ren_Direc1.Text = grd_Listad.Text
         
         grd_Listad.Col = 63
         txt_Ren_NomAr1.Text = grd_Listad.Text
         
         grd_Listad.Col = 64
         txt_Ren_Tele11.Text = grd_Listad.Text
         
         grd_Listad.Col = 65
         txt_Ren_Tele21.Text = grd_Listad.Text
         
         grd_Listad.Col = 66
         ipp_Ren_AlqMe1.Value = CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 67
         ipp_Ren_FIAlq1.Text = grd_Listad.Text
         
         grd_Listad.Col = 68
         If CInt(grd_Listad.Text) = 1 Then
            chk_Alqui2.Value = 1
            
            grd_Listad.Col = 69
            txt_Ren_Direc2.Text = grd_Listad.Text
            
            grd_Listad.Col = 70
            txt_Ren_NomAr2.Text = grd_Listad.Text
            
            grd_Listad.Col = 71
            txt_Ren_Tele12.Text = grd_Listad.Text
            
            grd_Listad.Col = 72
            txt_Ren_Tele22.Text = grd_Listad.Text
            
            grd_Listad.Col = 73
            ipp_Ren_AlqMe2.Value = CDbl(grd_Listad.Text)
            
            grd_Listad.Col = 74
            ipp_Ren_FIAlq2.Text = grd_Listad.Text
         End If
   
         grd_Listad.Col = 75
         If CInt(grd_Listad.Text) = 1 Then
            chk_Alqui3.Value = 1
            
            grd_Listad.Col = 76
            txt_Ren_Direc3.Text = grd_Listad.Text
            
            grd_Listad.Col = 77
            txt_Ren_NomAr3.Text = grd_Listad.Text
            
            grd_Listad.Col = 78
            txt_Ren_Tele13.Text = grd_Listad.Text
            
            grd_Listad.Col = 79
            txt_Ren_Tele23.Text = grd_Listad.Text
            
            grd_Listad.Col = 80
            ipp_Ren_AlqMe3.Value = CDbl(grd_Listad.Text)
            
            grd_Listad.Col = 81
            ipp_Ren_FIAlq3.Text = grd_Listad.Text
         End If
         
         grd_Listad.Col = 82
         ipp_Ren_IngNet.Value = CDbl(grd_Listad.Text)
   End Select
   
   Call gs_SetFocus(cmb_ActEco)
   
   cmd_Grabar.Enabled = False
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_Contad     As Integer
   Dim r_int_Linea1     As Integer
   Dim r_int_Linea2     As Integer
   Dim r_int_Linea3     As Integer
   
   r_int_Linea1 = 0
   r_int_Linea2 = 0
   r_int_Linea3 = 0
   
   If modatecli_g_int_Tip_ActEco = 1 Then
      'Inicializando Código de Actividad Económica - Titular
      
      modatecli_g_int_ActPri_Tit = 0
      modatecli_g_int_ActSec_Tit = 0
   Else
      'Inicializando Código de Actividad Económica - Cónyuge
      
      modatecli_g_int_ActPri_Cyg = 0
      modatecli_g_int_ActSec_Cyg = 0
   End If
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 2
      
      If r_int_Contad = 0 Then
         r_int_Linea1 = CInt(grd_Listad.Text)
      End If
      
      If r_int_Contad = 1 Then
         r_int_Linea2 = CInt(grd_Listad.Text)
      End If
   
      If r_int_Contad = 2 Then
         r_int_Linea3 = CInt(grd_Listad.Text)
      End If
      
      'Obteniendo Actividad Económica Principal
      If CInt(grd_Listad.Text) = 1 Then
         If modatecli_g_int_Tip_ActEco = 1 Then
            grd_Listad.Col = 3
            modatecli_g_int_ActPri_Tit = CInt(grd_Listad.Text)                   'Código de Actividad Económica
            
            grd_Listad.Col = 8
            modatecli_g_str_CodCiu_Tit = Format(CInt(grd_Listad.Text), "0000")   'Código CIIU
            
            grd_Listad.Col = 9
            modatecli_g_str_GirCom_Tit = grd_Listad.Text                         'Giro Comercial
            
            grd_Listad.Col = 27
            modatecli_g_str_SecEco_Tit = Format(CInt(grd_Listad.Text), "00")     'Sector Económico
            
            If modatecli_g_int_ActPri_Tit = 11 Or modatecli_g_int_ActPri_Tit = 31 Or modatecli_g_int_ActPri_Tit = 41 Then
               grd_Listad.Col = 4
               modatecli_g_int_TDoEmp_Tit = CInt(grd_Listad.Text)                'Tipo DOI Empresa
               
               grd_Listad.Col = 5
               modatecli_g_str_NDoEmp_Tit = grd_Listad.Text                      'Nro DOI Empresa
            Else
               modatecli_g_int_TDoEmp_Tit = 0
               modatecli_g_str_NDoEmp_Tit = ""
            End If
            
         Else
            grd_Listad.Col = 3
            modatecli_g_int_ActPri_Cyg = CInt(grd_Listad.Text)
         
            grd_Listad.Col = 8
            modatecli_g_str_CodCiu_Cyg = Format(CInt(grd_Listad.Text), "0000")   'Código CIIU
            
            grd_Listad.Col = 9
            modatecli_g_str_GirCom_Cyg = grd_Listad.Text                         'Giro Comercial
            
            grd_Listad.Col = 27
            modatecli_g_str_SecEco_Cyg = Format(CInt(grd_Listad.Text), "00")     'Sector Económico
            
            If modatecli_g_int_ActPri_Cyg = 11 Or modatecli_g_int_ActPri_Cyg = 31 Or modatecli_g_int_ActPri_Cyg = 41 Then
               grd_Listad.Col = 4
               modatecli_g_int_TDoEmp_Cyg = CInt(grd_Listad.Text)                'Tipo DOI Empresa
               
               grd_Listad.Col = 5
               modatecli_g_str_NDoEmp_Cyg = grd_Listad.Text                      'Nro DOI Empresa
            Else
               modatecli_g_int_TDoEmp_Cyg = 0
               modatecli_g_str_NDoEmp_Cyg = ""
            End If
         End If
      End If
      
      'Obteniendo Actividad Económica Secundaria
      grd_Listad.Col = 2
      If CInt(grd_Listad.Text) = 2 Then
         grd_Listad.Col = 3
         
         If modatecli_g_int_Tip_ActEco = 1 Then
            modatecli_g_int_ActSec_Tit = CInt(grd_Listad.Text)
         Else
            modatecli_g_int_ActSec_Cyg = CInt(grd_Listad.Text)
         End If
      End If
   Next r_int_Contad
   
   If r_int_Linea1 <> 1 And r_int_Linea2 <> 1 And r_int_Linea3 <> 1 Then
      MsgBox "Debe registar la Actividad Principal", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If (r_int_Linea1 = r_int_Linea2 And r_int_Linea1 <> 0 And r_int_Linea2 <> 0) Or (r_int_Linea1 = r_int_Linea3 And r_int_Linea1 <> 0 And r_int_Linea3 <> 0) Or (r_int_Linea2 = r_int_Linea3 And r_int_Linea2 <> 0 And r_int_Linea3 <> 0) Then
      MsgBox "No se puede tener dos Ordenes de Actividad igual.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If modatecli_g_int_Tip_ActEco = 1 Then
      Call fs_Graba_Arreglo_Tit
   Else
      Call fs_Graba_Arreglo_Cyg
   End If
   
   Unload Me
End Sub

Private Sub cmd_Ind_BusEmp_Click()
   If cmb_Ind_TDoEmp.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Ind_TDoEmp)
      Exit Sub
   End If
   
   If Len(Trim(txt_Ind_NDoEmp.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ind_NDoEmp)
      Exit Sub
   End If
   
   If Not gf_Valida_RUC(Mid(txt_Ind_NDoEmp.Text, 1, Len(txt_Ind_NDoEmp.Text) - 1), Right(txt_Ind_NDoEmp.Text, 1)) Then
      MsgBox "El Número de Documento de Identidad no es válido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ind_NDoEmp)
      Exit Sub
   End If
   
   Call fs_BusEmp(cmb_Ind_TDoEmp.ItemData(cmb_Ind_TDoEmp.ListIndex), txt_Ind_NDoEmp)
End Sub

Private Sub cmd_Ind_Direcc_Click()
   cmb_Ind_TipVia.ListIndex = frm_IngSol_01.cmb_TipVia.ListIndex
   txt_Ind_NomVia.Text = frm_IngSol_01.txt_NomVia.Text
   txt_Ind_Numero.Text = frm_IngSol_01.txt_Numero.Text
   txt_Ind_Interi.Text = frm_IngSol_01.txt_Interi.Text
   cmb_Ind_TipZon.ListIndex = frm_IngSol_01.cmb_TipZon.ListIndex
   txt_Ind_NomZon.Text = frm_IngSol_01.txt_NomZon.Text
   
   cmb_Ind_DptDir.ListIndex = frm_IngSol_01.cmb_DptDir.ListIndex
   
   Call moddat_gs_Carga_Provin(cmb_Ind_PrvDir, Format(cmb_Ind_DptDir.ItemData(cmb_Ind_DptDir.ListIndex), "00"))
   cmb_Ind_PrvDir.ListIndex = frm_IngSol_01.cmb_PrvDir.ListIndex
         
   Call moddat_gs_Carga_Distri(cmb_Ind_DstDir, Format(cmb_Ind_DptDir.ItemData(cmb_Ind_DptDir.ListIndex), "00"), Format(cmb_Ind_PrvDir.ItemData(cmb_Ind_PrvDir.ListIndex), "00"))
   cmb_Ind_DstDir.ListIndex = frm_IngSol_01.cmb_DstDir.ListIndex
   
   txt_Ind_Telef1.Text = frm_IngSol_01.txt_Telefo.Text
   
   Call gs_SetFocus(ipp_Ind_IngNet)
End Sub

Private Sub cmd_NueAct_Click()
   'Cargando Orden de Actividades Económicas
   If grd_Listad.Rows = 3 Then
      MsgBox "No puede ingresar más Actividades Económicas.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   l_int_FlgGrb = 1
   
   Call fs_Activa(True)
   Call fs_Activa_Dep(False)
   Call fs_Activa_Ind(False)
   Call fs_Activa_Com(False)
   Call fs_Activa_Acc(False)
   Call fs_Activa_Ren(False)
   
   Call gs_SetFocus(cmb_ActEco)
   
   cmd_Grabar.Enabled = False
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_int_TipIng  As Integer

   Dim r_int_Contad     As Integer
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt & " Ingreso de Solicitud de Crédito"
   
   'Cliente Titular
   If modatecli_g_int_Tip_ActEco = 1 Then
      pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
      
      r_int_TipIng = modatecli_g_int_ActEcoTit
   Else
      pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli & Chr(13) & Chr(10)
      pnl_Client.Caption = pnl_Client.Caption & CStr(modatecli_g_arr_DatGen(2).DatGen_TipDoc) & "-" & modatecli_g_arr_DatGen(2).DatGen_NumDoc & " / " & Trim(modatecli_g_arr_DatGen(2).DatGen_ApePat) & " " & Trim(modatecli_g_arr_DatGen(2).DatGen_ApeMat) & " " & Trim(modatecli_g_arr_DatGen(2).DatGen_Nombre)
      
      r_int_TipIng = modatecli_g_int_ActEcoCyg
   End If
   
   Call fs_Inicia
   Call gs_LimpiaGrid(grd_Listad)

   pnl_TraDep.Visible = True
   pnl_TraInd.Visible = False
   pnl_TraCom.Visible = False
   pnl_TraAcc.Visible = False
   pnl_TraRen.Visible = False

   Call fs_Limpia

   Call fs_Activa(False)
   Call fs_Activa_Dep(False)
   
   If r_int_TipIng = 1 Then
      cmd_Grabar.Enabled = False
      cmd_BorAct.Enabled = False
      cmd_EdiAct.Enabled = False
   ElseIf r_int_TipIng = 2 Then    'Si ya hay datos ingresados
      If modatecli_g_int_Tip_ActEco = 1 Then
         Call fs_Carga_Arreglo_Tit
      Else
         Call fs_Carga_Arreglo_Cyg
      End If
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad  As Integer
   
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 3215
   grd_Listad.ColWidth(1) = 5930
   
   For r_int_Contad = 2 To grd_Listad.Cols - 1
      grd_Listad.ColWidth(r_int_Contad) = 0
   Next r_int_Contad
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   Call moddat_gs_Carga_TipDocIde(cmb_Dep_TipDoc, 2)
   Call moddat_gs_Carga_TipDocIde(cmb_Ind_TipDoc, 2)
   Call moddat_gs_Carga_TipDocIde(cmb_Com_TipDoc, 2)
   Call moddat_gs_Carga_TipDocIde(cmb_Acc_TipDoc, 2)
   Call moddat_gs_Carga_TipDocIde(cmb_Ren_TipDoc, 2)
   Call moddat_gs_Carga_TipDocIde(cmb_Ind_TDoEmp, 2)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_ActEco, 1, "008")
   Call moddat_gs_Carga_LisIte_Combo(cmb_OrdAct, 1, "007")

   Call moddat_gs_Carga_LisIte_Combo(cmb_Dep_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Ind_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Com_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Acc_TipVia, 1, "201")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Dep_TipZon, 1, "202")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Ind_TipZon, 1, "202")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Com_TipZon, 1, "202")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Acc_TipZon, 1, "202")
   
   Call moddat_gs_Carga_Depart(cmb_Dep_DptDir)
   Call moddat_gs_Carga_Depart(cmb_Ind_DptDir)
   Call moddat_gs_Carga_Depart(cmb_Com_DptDir)
   Call moddat_gs_Carga_Depart(cmb_Acc_DptDir)
   
   Call moddat_gs_Carga_GirCom(cmb_Dep_GirCom, l_arr_Dep_GirCom)
   Call moddat_gs_Carga_GirCom(cmb_Ind_GirCom, l_arr_Ind_GirCom)
   Call moddat_gs_Carga_GirCom(cmb_Com_GirCom, l_arr_Com_GirCom)
   Call moddat_gs_Carga_GirCom(cmb_Acc_GirCom, l_arr_Acc_GirCom)
   Call moddat_gs_Carga_GirCom(cmb_Ren_GirCom, l_arr_Ren_GirCom)
   
   Call moddat_gs_Carga_LisIte(cmb_Dep_NomCar, l_arr_Dep_NomCar, 1, "503")
   Call moddat_gs_Carga_LisIte(cmb_Ind_NomCar, l_arr_Ind_NomCar, 1, "503")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Dep_FreHab, 1, "210")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Ind_ConLoc, 1, "214")
   Call moddat_gs_Carga_TipDocIde(cmb_Ind_TipDoc, 2)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Com_RegTri, 1, "215")
   'Call moddat_gs_Carga_TipViv_ResAct(cmb_Com_TipLoc, 3)
End Sub

Private Sub fs_Limpia_Dep()
   cmb_Dep_TipDoc.ListIndex = -1
   txt_Dep_NumDoc.Text = ""
   
   pnl_Dep_FlgEmp.Caption = ""
   
   txt_Dep_RazSoc.Text = ""
   txt_Dep_NomCom.Text = ""
   cmb_Dep_GirCom.ListIndex = -1
   txt_Dep_GirCom.Text = ""
   txt_Dep_GirCom.Enabled = False
   chk_Dep_Sucurs.Value = 0
   txt_Dep_Sucurs.Text = ""
   txt_Dep_Sucurs.Enabled = False
   cmb_Dep_TipVia.ListIndex = -1
   txt_Dep_NomVia.Text = ""
   txt_Dep_Numero.Text = ""
   txt_Dep_Interi.Text = ""
   cmb_Dep_TipZon.ListIndex = -1
   txt_Dep_NomZon.Text = ""
   cmb_Dep_DptDir.ListIndex = -1
   cmb_Dep_PrvDir.Clear
   cmb_Dep_DstDir.Clear
   txt_Dep_Refere.Text = ""
   txt_Dep_Telef1.Text = ""
   txt_Dep_Telef2.Text = ""
   txt_Dep_NumFax.Text = ""
   txt_Dep_TeleRH.Text = ""
   txt_Dep_AnexRH.Text = ""
   
   ipp_Dep_IngNet.Value = 0
   cmb_Dep_FreHab.ListIndex = -1
   cmb_Dep_NomCar.ListIndex = -1
   txt_Dep_NomCar.Text = ""
   txt_Dep_NomCar.Enabled = False
   txt_Dep_NomAre.Text = ""
   ipp_Dep_FecIng.Text = Format(CDate(CDate(moddat_g_str_FecSis) - CDate(365)), "dd/mm/yyyy")
   txt_Dep_NumAnx.Text = ""
   txt_Dep_TelDir.Text = ""
   txt_Dep_Celula.Text = ""
   txt_Dep_DirEle.Text = ""
   ipp_Dep_FecCes.Text = Format(CDate(CDate(moddat_g_str_FecSis) - CDate(365)), "dd/mm/yyyy")
   ipp_Dep_FecCes.Enabled = False
End Sub

Private Sub fs_Limpia_Ind()
   cmb_Ind_TipDoc.ListIndex = -1
   txt_Ind_NumDoc.Text = ""
   cmb_Ind_GirCom.ListIndex = -1
   txt_Ind_GirCom.Text = ""
   txt_Ind_GirCom.Enabled = False
   cmb_Ind_TipVia.ListIndex = -1
   txt_Ind_NomVia.Text = ""
   txt_Ind_Numero.Text = ""
   txt_Ind_Interi.Text = ""
   cmb_Ind_TipZon.ListIndex = -1
   txt_Ind_NomZon.Text = ""
   cmb_Ind_DptDir.ListIndex = -1
   cmb_Ind_PrvDir.Clear
   cmb_Ind_DstDir.Clear
   txt_Ind_Refere.Text = ""
   txt_Ind_Telef1.Text = ""
   txt_Ind_Telef2.Text = ""
   txt_Ind_NumFax.Text = ""
   
   ipp_Ind_IngNet.Value = 0
   cmb_Ind_ConLoc.ListIndex = -1
   cmb_Ind_TDoEmp.ListIndex = -1
   txt_Ind_NDoEmp.Text = ""
   
   cmd_Ind_BusEmp.Enabled = False
   pnl_Ind_FlgEmp.Caption = ""
   
   txt_Ind_RazSoc.Text = ""
   txt_Ind_Tl1Emp.Text = ""
   txt_Ind_Tl2Emp.Text = ""
   
   cmb_Ind_NomCar.ListIndex = -1
   txt_Ind_NomCar.Text = ""
   ipp_Ind_FecIng.Text = Format(CDate(CDate(moddat_g_str_FecSis) - CDate(365)), "dd/mm/yyyy")
End Sub

Private Sub fs_Limpia_Acc()
   cmb_Acc_TipDoc.ListIndex = -1
   txt_Acc_NumDoc.Text = ""
   
   pnl_Acc_FlgEmp.Caption = ""
   
   txt_Acc_RazSoc.Text = ""
   txt_Acc_NomCom.Text = ""
   cmb_Acc_GirCom.ListIndex = -1
   txt_Acc_GirCom.Enabled = False
   txt_Acc_GirCom.Text = ""
   cmb_Acc_TipVia.ListIndex = -1
   txt_Acc_NomVia.Text = ""
   txt_Acc_Numero.Text = ""
   txt_Acc_Interi.Text = ""
   cmb_Acc_TipZon.ListIndex = -1
   txt_Acc_NomZon.Text = ""
   cmb_Acc_DptDir.ListIndex = -1
   cmb_Acc_PrvDir.Clear
   cmb_Acc_DstDir.Clear
   txt_Acc_Refere.Text = ""
   txt_Acc_Telef1.Text = ""
   txt_Acc_Telef2.Text = ""
   txt_Acc_NumFax.Text = ""
   
   ipp_Acc_IngNet.Value = 0
   ipp_Acc_PorAcc.Value = 0
   ipp_Acc_FecAnt.Text = Format(CDate(CDate(moddat_g_str_FecSis) - CDate(365)), "dd/mm/yyyy")
End Sub

Private Sub fs_Limpia_Ren()
   cmb_Ren_TipDoc.ListIndex = -1
   txt_Ren_NumDoc.Text = ""
   cmb_Ren_GirCom.ListIndex = -1
   
   txt_Ren_Direc1.Text = ""
   txt_Ren_NomAr1.Text = ""
   txt_Ren_Tele11.Text = ""
   txt_Ren_Tele21.Text = ""
   ipp_Ren_AlqMe1.Value = 0
   ipp_Ren_FIAlq1.Text = Format(CDate(CDate(moddat_g_str_FecSis) - CDate(365)), "dd/mm/yyyy")
   
   chk_Alqui2.Value = 0
   txt_Ren_Direc2.Text = ""
   txt_Ren_NomAr2.Text = ""
   txt_Ren_Tele12.Text = ""
   txt_Ren_Tele22.Text = ""
   ipp_Ren_AlqMe2.Value = 0
   ipp_Ren_FIAlq2.Text = Format(CDate(CDate(moddat_g_str_FecSis) - CDate(365)), "dd/mm/yyyy")
   
   txt_Ren_Direc2.Enabled = False
   txt_Ren_NomAr2.Enabled = False
   txt_Ren_Tele12.Enabled = False
   txt_Ren_Tele22.Enabled = False
   ipp_Ren_AlqMe2.Enabled = False
   ipp_Ren_FIAlq2.Enabled = False
   
   
   chk_Alqui3.Value = 0
   txt_Ren_Direc3.Text = ""
   txt_Ren_NomAr3.Text = ""
   txt_Ren_Tele13.Text = ""
   txt_Ren_Tele23.Text = ""
   ipp_Ren_AlqMe3.Value = 0
   ipp_Ren_FIAlq3.Text = Format(CDate(CDate(moddat_g_str_FecSis) - CDate(365)), "dd/mm/yyyy")
   
   txt_Ren_Direc3.Enabled = False
   txt_Ren_NomAr3.Enabled = False
   txt_Ren_Tele13.Enabled = False
   txt_Ren_Tele23.Enabled = False
   ipp_Ren_AlqMe3.Enabled = False
   ipp_Ren_FIAlq3.Enabled = False
   
   ipp_Ren_IngNet.Value = 0
End Sub

Private Sub fs_Limpia_Com()
   cmb_Com_TipDoc.ListIndex = -1
   txt_Com_NumDoc.Text = ""
   
   pnl_Com_FlgEmp.Caption = ""
   
   txt_Com_RazSoc.Text = ""
   txt_Com_NomCom.Text = ""
   
   cmb_Com_GirCom.ListIndex = -1
   txt_Com_GirCom.Text = ""
   txt_Com_GirCom.Enabled = False
   cmb_Com_TipVia.ListIndex = -1
   txt_Com_NomVia.Text = ""
   txt_Com_Numero.Text = ""
   txt_Com_Interi.Text = ""
   cmb_Com_TipZon.ListIndex = -1
   txt_Com_NomZon.Text = ""
   cmb_Com_DptDir.ListIndex = -1
   cmb_Com_PrvDir.Clear
   cmb_Com_DstDir.Clear
   txt_Com_Refere.Text = ""
   txt_Com_Telef1.Text = ""
   txt_Com_Telef2.Text = ""
   txt_Com_NumFax.Text = ""
   
   ipp_Com_IngNet.Value = 0
   ipp_Com_VtaMen.Value = 0
   
   ipp_Com_FecIni.Text = Format(CDate(CDate(moddat_g_str_FecSis) - CDate(365)), "dd/mm/yyyy")
   
   cmb_Com_RegTri.ListIndex = -1
   ipp_Com_PorPar.Value = 0
   cmb_Com_TipLoc.ListIndex = -1
   ipp_Com_AlqMen.Value = 0
   txt_Com_NomArr.Text = ""
   txt_Com_Tl1Arr.Text = ""
   txt_Com_Tl2Arr.Text = ""

   ipp_Com_AlqMen.Enabled = False
   txt_Com_NomArr.Enabled = False
   txt_Com_Tl1Arr.Enabled = False
   txt_Com_Tl2Arr.Enabled = False
End Sub

Private Sub cmb_Dep_DptDir_Click()
   If cmb_Dep_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Dep_PrvDir.Clear
         cmb_Dep_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Dep_PrvDir, Format(cmb_Dep_DptDir.ItemData(cmb_Dep_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Dep_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_Dep_DptDir_Change()
   l_str_Dep_DptDir = cmb_Dep_DptDir.Text
End Sub

Private Sub cmb_Dep_DptDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Dep_DptDir = cmb_Dep_DptDir.Text
End Sub

Private Sub cmb_Dep_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Dep_DptDir, l_str_Dep_DptDir)
      l_int_FlgCmb = True
      
      cmb_Dep_PrvDir.Clear
      cmb_Dep_DstDir.Clear
      
      If cmb_Dep_DptDir.ListIndex > -1 Then
         l_str_Dep_DptDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Dep_PrvDir, Format(cmb_Dep_DptDir.ItemData(cmb_Dep_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Dep_PrvDir)
   End If
End Sub

Private Sub cmb_Dep_PrvDir_Change()
   l_str_Dep_PrvDir = cmb_Dep_PrvDir.Text
End Sub

Private Sub cmb_Dep_PrvDir_Click()
   If cmb_Dep_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Dep_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Dep_DstDir, Format(cmb_Dep_DptDir.ItemData(cmb_Dep_DptDir.ListIndex), "00"), Format(cmb_Dep_PrvDir.ItemData(cmb_Dep_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Dep_DstDir)
      End If
   End If
End Sub

Private Sub cmb_Dep_PrvDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Dep_PrvDir = cmb_Dep_PrvDir.Text
End Sub

Private Sub cmb_Dep_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Dep_PrvDir, l_str_Dep_PrvDir)
      l_int_FlgCmb = True
      
      cmb_Dep_DstDir.Clear
      If cmb_Dep_PrvDir.ListIndex > -1 Then
         l_str_Dep_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Dep_DstDir, Format(cmb_Dep_DptDir.ItemData(cmb_Dep_DptDir.ListIndex), "00"), Format(cmb_Dep_PrvDir.ItemData(cmb_Dep_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Dep_DstDir)
   End If
End Sub

Private Sub cmb_Dep_DstDir_Change()
   l_str_Dep_DstDir = cmb_Dep_DstDir.Text
End Sub

Private Sub cmb_Dep_DstDir_Click()
   If cmb_Dep_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Dep_Refere)
      End If
   End If
End Sub

Private Sub cmb_Dep_DstDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Dep_DstDir = cmb_Dep_DstDir.Text
End Sub

Private Sub cmb_Dep_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Dep_DstDir, l_str_Dep_DstDir)
      l_int_FlgCmb = True
      
      If cmb_Dep_DstDir.ListIndex > -1 Then
         l_str_Dep_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Dep_Refere)
   End If
End Sub

Private Sub cmb_Ind_DptDir_Click()
   If cmb_Ind_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Ind_PrvDir.Clear
         cmb_Ind_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Ind_PrvDir, Format(cmb_Ind_DptDir.ItemData(cmb_Ind_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Ind_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_Ind_DptDir_Change()
   l_str_Ind_DptDir = cmb_Ind_DptDir.Text
End Sub

Private Sub cmb_Ind_DptDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Ind_DptDir = cmb_Ind_DptDir.Text
End Sub

Private Sub cmb_Ind_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Ind_DptDir, l_str_Ind_DptDir)
      l_int_FlgCmb = True
      
      cmb_Ind_PrvDir.Clear
      cmb_Ind_DstDir.Clear
      
      If cmb_Ind_DptDir.ListIndex > -1 Then
         l_str_Ind_DptDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Ind_PrvDir, Format(cmb_Ind_DptDir.ItemData(cmb_Ind_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Ind_PrvDir)
   End If
End Sub

Private Sub cmb_Ind_PrvDir_Change()
   l_str_Ind_PrvDir = cmb_Ind_PrvDir.Text
End Sub

Private Sub cmb_Ind_PrvDir_Click()
   If cmb_Ind_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Ind_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Ind_DstDir, Format(cmb_Ind_DptDir.ItemData(cmb_Ind_DptDir.ListIndex), "00"), Format(cmb_Ind_PrvDir.ItemData(cmb_Ind_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Ind_DstDir)
      End If
   End If
End Sub

Private Sub cmb_Ind_PrvDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Ind_PrvDir = cmb_Ind_PrvDir.Text
End Sub

Private Sub cmb_Ind_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Ind_PrvDir, l_str_Ind_PrvDir)
      l_int_FlgCmb = True
      
      cmb_Ind_DstDir.Clear
      If cmb_Ind_PrvDir.ListIndex > -1 Then
         l_str_Ind_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Ind_DstDir, Format(cmb_Ind_DptDir.ItemData(cmb_Ind_DptDir.ListIndex), "00"), Format(cmb_Ind_PrvDir.ItemData(cmb_Ind_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Ind_DstDir)
   End If
End Sub

Private Sub cmb_Ind_DstDir_Change()
   l_str_Ind_DstDir = cmb_Ind_DstDir.Text
End Sub

Private Sub cmb_Ind_DstDir_Click()
   If cmb_Ind_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Ind_Refere)
      End If
   End If
End Sub

Private Sub cmb_Ind_DstDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Ind_DstDir = cmb_Ind_DstDir.Text
End Sub

Private Sub cmb_Ind_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Ind_DstDir, l_str_Ind_DstDir)
      l_int_FlgCmb = True
      
      If cmb_Ind_DstDir.ListIndex > -1 Then
         l_str_Ind_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Ind_Refere)
   End If
End Sub

Private Sub cmb_Com_DptDir_Click()
   If cmb_Com_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Com_PrvDir.Clear
         cmb_Com_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Com_PrvDir, Format(cmb_Com_DptDir.ItemData(cmb_Com_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Com_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_Com_DptDir_Change()
   l_str_Com_DptDir = cmb_Com_DptDir.Text
End Sub

Private Sub cmb_Com_DptDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Com_DptDir = cmb_Com_DptDir.Text
End Sub

Private Sub cmb_Com_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Com_DptDir, l_str_Com_DptDir)
      l_int_FlgCmb = True
      
      cmb_Com_PrvDir.Clear
      cmb_Com_DstDir.Clear
      
      If cmb_Com_DptDir.ListIndex > -1 Then
         l_str_Com_DptDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Com_PrvDir, Format(cmb_Com_DptDir.ItemData(cmb_Com_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Com_PrvDir)
   End If
End Sub

Private Sub cmb_Com_PrvDir_Change()
   l_str_Com_PrvDir = cmb_Com_PrvDir.Text
End Sub

Private Sub cmb_Com_PrvDir_Click()
   If cmb_Com_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Com_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Com_DstDir, Format(cmb_Com_DptDir.ItemData(cmb_Com_DptDir.ListIndex), "00"), Format(cmb_Com_PrvDir.ItemData(cmb_Com_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Com_DstDir)
      End If
   End If
End Sub

Private Sub cmb_Com_PrvDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Com_PrvDir = cmb_Com_PrvDir.Text
End Sub

Private Sub cmb_Com_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Com_PrvDir, l_str_Com_PrvDir)
      l_int_FlgCmb = True
      
      cmb_Com_DstDir.Clear
      If cmb_Com_PrvDir.ListIndex > -1 Then
         l_str_Com_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Com_DstDir, Format(cmb_Com_DptDir.ItemData(cmb_Com_DptDir.ListIndex), "00"), Format(cmb_Com_PrvDir.ItemData(cmb_Com_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Com_DstDir)
   End If
End Sub

Private Sub cmb_Com_DstDir_Change()
   l_str_Com_DstDir = cmb_Com_DstDir.Text
End Sub

Private Sub cmb_Com_DstDir_Click()
   If cmb_Com_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Com_Refere)
      End If
   End If
End Sub

Private Sub cmb_Com_DstDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Com_DstDir = cmb_Com_DstDir.Text
End Sub

Private Sub cmb_Com_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Com_DstDir, l_str_Com_DstDir)
      l_int_FlgCmb = True
      
      If cmb_Com_DstDir.ListIndex > -1 Then
         l_str_Com_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Com_Refere)
   End If
End Sub

Private Sub cmb_Acc_DptDir_Click()
   If cmb_Acc_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Acc_PrvDir.Clear
         cmb_Acc_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Acc_PrvDir, Format(cmb_Acc_DptDir.ItemData(cmb_Acc_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Acc_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_Acc_DptDir_Change()
   l_str_Acc_DptDir = cmb_Acc_DptDir.Text
End Sub

Private Sub cmb_Acc_DptDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Acc_DptDir = cmb_Acc_DptDir.Text
End Sub

Private Sub cmb_Acc_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Acc_DptDir, l_str_Acc_DptDir)
      l_int_FlgCmb = True
      
      cmb_Acc_PrvDir.Clear
      cmb_Acc_DstDir.Clear
      
      If cmb_Acc_DptDir.ListIndex > -1 Then
         l_str_Acc_DptDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Acc_PrvDir, Format(cmb_Acc_DptDir.ItemData(cmb_Acc_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Acc_PrvDir)
   End If
End Sub

Private Sub cmb_Acc_PrvDir_Change()
   l_str_Acc_PrvDir = cmb_Acc_PrvDir.Text
End Sub

Private Sub cmb_Acc_PrvDir_Click()
   If cmb_Acc_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Acc_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Acc_DstDir, Format(cmb_Acc_DptDir.ItemData(cmb_Acc_DptDir.ListIndex), "00"), Format(cmb_Acc_PrvDir.ItemData(cmb_Acc_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Acc_DstDir)
      End If
   End If
End Sub

Private Sub cmb_Acc_PrvDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Acc_PrvDir = cmb_Acc_PrvDir.Text
End Sub

Private Sub cmb_Acc_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Acc_PrvDir, l_str_Acc_PrvDir)
      l_int_FlgCmb = True
      
      cmb_Acc_DstDir.Clear
      If cmb_Acc_PrvDir.ListIndex > -1 Then
         l_str_Acc_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Acc_DstDir, Format(cmb_Acc_DptDir.ItemData(cmb_Acc_DptDir.ListIndex), "00"), Format(cmb_Acc_PrvDir.ItemData(cmb_Acc_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Acc_DstDir)
   End If
End Sub

Private Sub cmb_Acc_DstDir_Change()
   l_str_Acc_DstDir = cmb_Acc_DstDir.Text
End Sub

Private Sub cmb_Acc_DstDir_Click()
   If cmb_Acc_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Acc_Refere)
      End If
   End If
End Sub

Private Sub cmb_Acc_DstDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Acc_DstDir = cmb_Acc_DstDir.Text
End Sub

Private Sub cmb_Acc_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Acc_DstDir, l_str_Acc_DstDir)
      l_int_FlgCmb = True
      
      If cmb_Acc_DstDir.ListIndex > -1 Then
         l_str_Acc_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Acc_Refere)
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_Acc_FecAnt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Agrega)
   End If
End Sub

Private Sub ipp_Acc_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Acc_PorAcc)
   End If
End Sub

Private Sub ipp_Acc_PorAcc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Acc_FecAnt)
   End If
End Sub

Private Sub ipp_Com_AlqMen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com_NomArr)
   End If
End Sub

Private Sub ipp_Com_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Com_RegTri)
   End If
End Sub

Private Sub ipp_Com_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Com_VtaMen)
   End If
End Sub

Private Sub ipp_Com_PorPar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Com_TipLoc)
   End If
End Sub

Private Sub ipp_Com_VtaMen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Com_FecIni)
   End If
End Sub

Private Sub ipp_Dep_FecIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_NumAnx)
   End If
End Sub

Private Sub ipp_Ind_FecIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Agrega)
   End If
End Sub

Private Sub ipp_Ind_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Ind_ConLoc)
   End If
End Sub

Private Sub ipp_Ind_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Ind_FecIni)
   End If
End Sub

Private Sub ipp_Ren_AlqMe1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Ren_FIAlq1)
   End If
End Sub

Private Sub ipp_Ren_FIAlq1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(chk_Alqui2)
   End If
End Sub

Private Sub ipp_Ren_AlqMe2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Ren_FIAlq2)
   End If
End Sub

Private Sub ipp_Ren_FIAlq2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(chk_Alqui3)
   End If
End Sub

Private Sub ipp_Ren_AlqMe3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Ren_FIAlq3)
   End If
End Sub

Private Sub ipp_Ren_FIAlq3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Ren_IngNet)
   End If
End Sub

Private Sub ipp_Ren_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Agrega)
   End If
End Sub

Private Sub txt_Acc_NumDoc_Change()
   pnl_Acc_FlgEmp.Caption = ""

   txt_Acc_RazSoc.Enabled = True
   txt_Acc_NomCom.Enabled = True
   cmb_Acc_GirCom.Enabled = True
   txt_Acc_GirCom.Enabled = True
   cmb_Acc_TipVia.Enabled = True
   txt_Acc_NomVia.Enabled = True
   cmb_Acc_TipZon.Enabled = True
   txt_Acc_NomZon.Enabled = True
   cmb_Acc_DptDir.Enabled = True
   cmb_Acc_PrvDir.Enabled = True
   cmb_Acc_DstDir.Enabled = True
   txt_Acc_Refere.Enabled = True
   txt_Acc_Telef1.Enabled = True
   txt_Acc_Telef2.Enabled = True
   txt_Acc_NumFax.Enabled = True
End Sub

Private Sub txt_Com_NumDoc_Change()
   pnl_Com_FlgEmp.Caption = ""
   
   txt_Com_RazSoc.Enabled = True
   txt_Com_NomCom.Enabled = True
   cmb_Com_GirCom.Enabled = True
   txt_Com_GirCom.Enabled = True
   cmb_Com_TipVia.Enabled = True
   txt_Com_NomVia.Enabled = True
   cmb_Com_TipZon.Enabled = True
   txt_Com_NomZon.Enabled = True
   cmb_Com_DptDir.Enabled = True
   cmb_Com_PrvDir.Enabled = True
   cmb_Com_DstDir.Enabled = True
   txt_Com_Refere.Enabled = True
   txt_Com_Telef1.Enabled = True
   txt_Com_Telef2.Enabled = True
   txt_Com_NumFax.Enabled = True
End Sub

Private Sub txt_Dep_NumDoc_Change()
   pnl_Dep_FlgEmp.Caption = ""
   
   txt_Dep_RazSoc.Enabled = True
   txt_Dep_NomCom.Enabled = True
   cmb_Dep_GirCom.Enabled = True
   txt_Dep_GirCom.Enabled = True
   cmb_Dep_TipVia.Enabled = True
   txt_Dep_NomVia.Enabled = True
   cmb_Dep_TipZon.Enabled = True
   txt_Dep_NomZon.Enabled = True
   cmb_Dep_DptDir.Enabled = True
   cmb_Dep_PrvDir.Enabled = True
   cmb_Dep_DstDir.Enabled = True
   txt_Dep_Refere.Enabled = True
   txt_Dep_Telef1.Enabled = True
   txt_Dep_Telef2.Enabled = True
   txt_Dep_NumFax.Enabled = True
   txt_Dep_TeleRH.Enabled = True
   txt_Dep_AnexRH.Enabled = True
End Sub

Private Sub txt_Dep_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_Dep_NumDoc)
End Sub

Private Sub txt_Dep_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Dep_BusEmp)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Dep_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_Dep_RazSoc)
End Sub

Private Sub txt_Dep_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_NomCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_Dep_NomCom_GotFocus()
   Call gs_SelecTodo(txt_Dep_NomCom)
End Sub

Private Sub txt_Dep_NomCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Dep_GirCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub cmb_Dep_GirCom_Change()
   l_str_Dep_GirCom = cmb_Dep_GirCom.Text
End Sub

Private Sub cmb_Dep_GirCom_Click()
   txt_Dep_GirCom.Enabled = False
   txt_Dep_GirCom.Text = ""
   
   If cmb_Dep_GirCom.ListIndex > -1 Then
      If l_int_FlgCmb Then
         If l_arr_Dep_GirCom(cmb_Dep_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
            txt_Dep_GirCom.Enabled = True
            Call gs_SetFocus(txt_Dep_GirCom)
         Else
            Call gs_SetFocus(cmb_Dep_TipVia)
         End If
      End If
   End If
End Sub

Private Sub cmb_Dep_GirCom_GotFocus()
   l_int_FlgCmb = True
   l_str_Dep_GirCom = cmb_Dep_GirCom.Text
End Sub

Private Sub cmb_Dep_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ./*+#,()" + Chr(34))
   Else
      txt_Dep_GirCom.Enabled = False
      txt_Dep_GirCom.Text = ""
      
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Dep_GirCom, l_str_Dep_GirCom)
      l_int_FlgCmb = True
      
      If cmb_Dep_GirCom.ListIndex > -1 Then
         l_str_Dep_GirCom = ""
      End If
      
      If l_arr_Dep_GirCom(cmb_Dep_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
         txt_Dep_GirCom.Enabled = True
         Call gs_SetFocus(txt_Dep_GirCom)
      Else
         Call gs_SetFocus(cmb_Dep_TipVia)
      End If
   End If
End Sub

Private Sub txt_Dep_GirCom_GotFocus()
   Call gs_SelecTodo(txt_Dep_GirCom)
End Sub

Private Sub txt_Dep_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(chk_Dep_Sucurs)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_Dep_Sucurs_GotFocus()
   Call gs_SelecTodo(txt_Dep_Sucurs)
End Sub

Private Sub txt_Dep_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Dep_TipVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub cmb_Dep_TipVia_Click()
   Call gs_SetFocus(txt_Dep_NomVia)
End Sub

Private Sub cmb_Dep_TipVia_KeyPress(KeyAscii As Integer)
   Call cmb_Dep_TipVia_Click
End Sub

Private Sub txt_Dep_NomVia_GotFocus()
   Call gs_SelecTodo(txt_Dep_NomVia)
End Sub

Private Sub txt_Dep_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_Numero)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Dep_Numero_GotFocus()
   Call gs_SelecTodo(txt_Dep_Numero)
End Sub

Private Sub txt_Dep_Numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_Interi)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Dep_Interi_GotFocus()
   Call gs_SelecTodo(txt_Dep_Interi)
End Sub

Private Sub txt_Dep_Interi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Dep_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub cmb_Dep_TipZon_Click()
   Call gs_SetFocus(txt_Dep_NomZon)
End Sub

Private Sub cmb_Dep_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_Dep_TipZon_Click
End Sub

Private Sub txt_Dep_NomZon_GotFocus()
   Call gs_SelecTodo(txt_Dep_NomZon)
End Sub

Private Sub txt_Dep_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Dep_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Dep_Refere_GotFocus()
   Call gs_SelecTodo(txt_Dep_Refere)
End Sub

Private Sub txt_Dep_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Dep_Telef1_GotFocus()
   Call gs_SelecTodo(txt_Dep_Telef1)
End Sub

Private Sub txt_Dep_Telef1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_Telef2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Dep_Telef2_GotFocus()
   Call gs_SelecTodo(txt_Dep_Telef2)
End Sub

Private Sub txt_Dep_Telef2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_NumFax)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Dep_NumFax_GotFocus()
   Call gs_SelecTodo(txt_Dep_NumFax)
End Sub

Private Sub txt_Dep_NumFax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_TeleRH)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Dep_TeleRH_GotFocus()
   Call gs_SelecTodo(txt_Dep_TeleRH)
End Sub

Private Sub txt_Dep_TeleRH_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_AnexRH)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Dep_AnexRH_GotFocus()
   Call gs_SelecTodo(txt_Dep_AnexRH)
End Sub

Private Sub txt_Dep_AnexRH_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Dep_IngNet)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub ipp_Dep_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Dep_FreHab)
   End If
End Sub

Private Sub cmb_Dep_NomCar_Change()
   l_str_Dep_NomCar = cmb_Dep_NomCar.Text
End Sub

Private Sub cmb_Dep_NomCar_Click()
   txt_Dep_NomCar.Enabled = False
   txt_Dep_NomCar.Text = ""
   
   If cmb_Dep_NomCar.ListIndex > -1 Then
      If l_int_FlgCmb Then
         If l_arr_Dep_NomCar(cmb_Dep_NomCar.ListIndex + 1).Genera_Codigo = "999999" Then
            txt_Dep_NomCar.Enabled = True
            Call gs_SetFocus(txt_Dep_NomCar)
         Else
            Call gs_SetFocus(txt_Dep_NomAre)
         End If
      End If
   End If
End Sub

Private Sub cmb_Dep_NomCar_GotFocus()
   l_int_FlgCmb = True
   l_str_Dep_NomCar = cmb_Dep_NomCar.Text
End Sub

Private Sub cmb_Dep_NomCar_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ./*+#,()<>" + Chr(34))
   Else
      txt_Dep_NomCar.Enabled = False
      txt_Dep_NomCar.Text = ""
      
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Dep_NomCar, l_str_Dep_NomCar)
      l_int_FlgCmb = True
      
      If cmb_Dep_NomCar.ListIndex > -1 Then
         l_str_Dep_NomCar = ""
      End If
      
      If l_arr_Dep_NomCar(cmb_Dep_NomCar.ListIndex + 1).Genera_Codigo = "999999" Then
         txt_Dep_NomCar.Enabled = True
         Call gs_SetFocus(txt_Dep_NomCar)
      Else
         Call gs_SetFocus(txt_Dep_NomAre)
      End If
   End If
End Sub

Private Sub txt_Dep_NomCar_GotFocus()
   Call gs_SelecTodo(txt_Dep_NomCar)
End Sub

Private Sub txt_Dep_NomCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_NomAre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_Dep_NomAre_GotFocus()
   Call gs_SelecTodo(txt_Dep_NomAre)
End Sub

Private Sub txt_Dep_NomAre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Dep_FecIng)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_Dep_NumAnx_GotFocus()
   Call gs_SelecTodo(txt_Dep_NumAnx)
End Sub

Private Sub txt_Dep_NumAnx_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_TelDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Dep_TelDir_GotFocus()
   Call gs_SelecTodo(txt_Dep_TelDir)
End Sub

Private Sub txt_Dep_TelDir_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_Celula)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Dep_Celula_GotFocus()
   Call gs_SelecTodo(txt_Dep_Celula)
End Sub

Private Sub txt_Dep_Celula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_DirEle)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Dep_DirEle_GotFocus()
   Call gs_SelecTodo(txt_Dep_DirEle)
End Sub

Private Sub txt_Dep_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_Dep_FecCes.Enabled Then
         Call gs_SetFocus(ipp_Dep_FecCes)
      Else
         Call gs_SetFocus(cmd_Agrega)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
   End If
End Sub

Private Function ff_Valida_Dep() As Integer
   ff_Valida_Dep = False
   
   If cmb_Dep_TipDoc.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Dep_TipDoc)
      Exit Function
   End If
   
   If Len(Trim(txt_Dep_NumDoc.Text)) <> 11 Then
      MsgBox "Ingrese correctamente el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Dep_NumDoc)
      Exit Function
   End If

   If Not gf_Valida_RUC(Mid(txt_Dep_NumDoc.Text, 1, Len(txt_Dep_NumDoc.Text) - 1), Right(txt_Dep_NumDoc.Text, 1)) Then
      MsgBox "El Número de Documento de Identidad no es válido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Dep_NumDoc)
      Exit Function
   End If
   
   If Len(Trim(pnl_Dep_FlgEmp.Caption)) = 0 Then
      MsgBox "Debe buscar la empresa en el Maestro de Empresas.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Dep_BusEmp)
      Exit Function
   End If
   
   If Len(Trim(txt_Dep_RazSoc.Text)) = 0 Then
      MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Dep_RazSoc)
      Exit Function
   End If

   If Len(Trim(txt_Dep_NomCom.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Dep_NomCom)
      Exit Function
   End If

   If cmb_Dep_GirCom.ListIndex = -1 Then
      MsgBox "Seleccione el Giro Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Dep_GirCom)
      Exit Function
   End If

   If txt_Dep_GirCom.Enabled Then
      If Len(Trim(txt_Dep_GirCom.Text)) = 0 Then
         MsgBox "Debe ingresar el Giro Comercial.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Dep_GirCom)
         Exit Function
      End If
   End If
   
   If cmb_Dep_TipVia.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Vía de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Dep_TipVia)
      Exit Function
   End If

   If Len(Trim(txt_Dep_NomVia.Text)) = 0 Then
      MsgBox "Ingrese el Nombre de la Vóa de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Dep_NomVia)
      Exit Function
   End If
   
   If Len(Trim(txt_Dep_Numero.Text)) = 0 Then
      MsgBox "Ingrese el Número en la Vía de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Dep_Numero)
      Exit Function
   End If
   
   If cmb_Dep_TipZon.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Zona de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Dep_TipZon)
      Exit Function
   End If

   If cmb_Dep_TipZon.ItemData(cmb_Dep_TipZon.ListIndex) <> 12 Then
      If Len(Trim(txt_Dep_NomZon.Text)) = 0 Then
         MsgBox "Ingrese el Nombre de la Zona de la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Dep_NomZon)
         Exit Function
      End If
   End If

   If cmb_Dep_DptDir.ListIndex = -1 Then
      MsgBox "Seleccione el Departamento de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Dep_DptDir)
      Exit Function
   End If
   
   If cmb_Dep_PrvDir.ListIndex = -1 Then
      MsgBox "Seleccione la Provincia de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Dep_PrvDir)
      Exit Function
   End If
   
   If cmb_Dep_DstDir.ListIndex = -1 Then
      MsgBox "Seleccione el Distrito de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Dep_DstDir)
      Exit Function
   End If
   
   If Len(Trim(txt_Dep_Telef1.Text)) = 0 Then
      MsgBox "Ingrese el Teléfono de la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Dep_Telef1)
      Exit Function
   End If
   
   If Len(Trim(txt_Dep_TeleRH.Text)) = 0 And Len(Trim(txt_Dep_AnexRH.Text)) = 0 Then
      MsgBox "Ingrese el Teléfono o el Anexo de Recuros Humanos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Dep_TeleRH)
      Exit Function
   End If
   
   If CDbl(ipp_Dep_IngNet.Text) = 0 Then
      MsgBox "Ingrese el Ingreso Neto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Dep_IngNet)
      Exit Function
   End If
   
   If cmb_Dep_FreHab.ListIndex = -1 Then
      MsgBox "Seleccione la Frecuencia de Haberes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Dep_FreHab)
      Exit Function
   End If
   
   If cmb_Dep_NomCar.ListIndex = -1 Then
      MsgBox "Seleccione el Cargo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Dep_NomCar)
      Exit Function
   End If
   
   If txt_Dep_NomCar.Enabled Then
      If Len(Trim(txt_Dep_NomCar.Text)) = 0 Then
         MsgBox "Debe ingresar el Cargo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Dep_NomCar)
         Exit Function
      End If
   End If
   
   If Len(Trim(txt_Dep_NomAre.Text)) = 0 Then
      MsgBox "Debe ingresar el Area.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Dep_NomAre)
      Exit Function
   End If
   
   'Validar Fecha de Ingreso contra Matrices de Credito
   
   
   ff_Valida_Dep = True
End Function

Private Function ff_Valida_Com() As Integer
   ff_Valida_Com = False
   
   If cmb_Com_TipDoc.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Com_TipDoc)
      Exit Function
   End If
   
   If Len(Trim(txt_Com_NumDoc.Text)) <> 11 Then
      MsgBox "Ingrese correctamente el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Com_NumDoc)
      Exit Function
   End If

   If Not gf_Valida_RUC(Mid(txt_Com_NumDoc.Text, 1, Len(txt_Com_NumDoc.Text) - 1), Right(txt_Com_NumDoc.Text, 1)) Then
      MsgBox "El Número de Documento de Identidad no es válido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Com_NumDoc)
      Exit Function
   End If
   
   If Len(Trim(pnl_Com_FlgEmp.Caption)) = 0 Then
      MsgBox "Debe buscar la empresa en el Maestro de Empresas.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Com_BusEmp)
      Exit Function
   End If
   
   If Len(Trim(txt_Com_RazSoc.Text)) = 0 Then
      MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Com_RazSoc)
      Exit Function
   End If

   If Len(Trim(txt_Com_NomCom.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Com_NomCom)
      Exit Function
   End If
  
   If cmb_Com_GirCom.ListIndex = -1 Then
      MsgBox "Seleccione el Giro Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Com_GirCom)
      Exit Function
   End If

   If txt_Com_GirCom.Enabled Then
      If Len(Trim(txt_Com_GirCom.Text)) = 0 Then
         MsgBox "Debe ingresar el Giro Comercial.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Com_GirCom)
         Exit Function
      End If
   End If
   
   If cmb_Com_TipVia.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Vía de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Com_TipVia)
      Exit Function
   End If

   If Len(Trim(txt_Com_NomVia.Text)) = 0 Then
      MsgBox "Ingrese el Nombre de la Vóa de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Com_NomVia)
      Exit Function
   End If
   
   If Len(Trim(txt_Com_Numero.Text)) = 0 Then
      MsgBox "Ingrese el Número en la Vía de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Com_Numero)
      Exit Function
   End If
   
   If cmb_Com_TipZon.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Zona de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Com_TipZon)
      Exit Function
   End If

   If cmb_Com_TipZon.ItemData(cmb_Com_TipZon.ListIndex) <> 12 Then
      If Len(Trim(txt_Com_NomZon.Text)) = 0 Then
         MsgBox "Ingrese el Nombre de la Zona de la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Com_NomZon)
         Exit Function
      End If
   End If

   If cmb_Com_DptDir.ListIndex = -1 Then
      MsgBox "Seleccione el Departamento de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Com_DptDir)
      Exit Function
   End If
   
   If cmb_Com_PrvDir.ListIndex = -1 Then
      MsgBox "Seleccione la Provincia de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Com_PrvDir)
      Exit Function
   End If
   
   If cmb_Com_DstDir.ListIndex = -1 Then
      MsgBox "Seleccione el Distrito de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Com_DstDir)
      Exit Function
   End If
   
   If Len(Trim(txt_Com_Telef1.Text)) = 0 Then
      MsgBox "Ingrese el Teléfono de la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Com_Telef1)
      Exit Function
   End If
   
   If CDbl(ipp_Com_IngNet.Text) = 0 Then
      MsgBox "Ingrese el Ingreso Neto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Com_IngNet)
      Exit Function
   End If
   
   If CDbl(ipp_Com_VtaMen.Text) = 0 Then
      MsgBox "Ingrese las Ventas Mensuales.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Com_VtaMen)
      Exit Function
   End If
   
   'Validar Fecha de Inicio Operaciones
   
   If cmb_Com_RegTri.ListIndex = -1 Then
      MsgBox "Seleccione el Régimen Tributario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Com_RegTri)
      Exit Function
   End If
   
   If CDbl(ipp_Com_PorPar.Text) = 0 Then
      MsgBox "Ingrese el Porcentaje de Participación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Com_PorPar)
      Exit Function
   End If
   
   If cmb_Com_TipLoc.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Local Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Com_TipLoc)
      Exit Function
   End If
   
   If cmb_Com_TipDoc.ItemData(cmb_Com_TipDoc.ListIndex) = 2 Then
      If CDbl(ipp_Com_AlqMen.Text) = 0 Then
         MsgBox "Ingrese el Importe de Alquiler Mensual.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_Com_AlqMen)
         Exit Function
      End If
   
      If Len(Trim(txt_Com_NomArr.Text)) = 0 Then
         MsgBox "Ingrese el Nombre del Arrendador.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Com_NomArr)
         Exit Function
      End If
   
      If Len(Trim(txt_Com_Tl1Arr.Text)) = 0 Then
         MsgBox "Ingrese el Teléfono del Arrendador.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Com_Tl1Arr)
         Exit Function
      End If
   End If
   
   ff_Valida_Com = True
End Function

Private Function ff_Valida_Acc() As Integer
   ff_Valida_Acc = False
   
   If cmb_Acc_TipDoc.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Acc_TipDoc)
      Exit Function
   End If
   
   If Len(Trim(txt_Acc_NumDoc.Text)) <> 11 Then
      MsgBox "Ingrese correctamente el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Acc_NumDoc)
      Exit Function
   End If

   If Not gf_Valida_RUC(Mid(txt_Acc_NumDoc.Text, 1, Len(txt_Acc_NumDoc.Text) - 1), Right(txt_Acc_NumDoc.Text, 1)) Then
      MsgBox "El Número de Documento de Identidad no es válido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Acc_NumDoc)
      Exit Function
   End If
   
   If Len(Trim(pnl_Acc_FlgEmp.Caption)) = 0 Then
      MsgBox "Debe buscar la empresa en el Maestro de Empresas.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Acc_BusEmp)
      Exit Function
   End If
   
   If Len(Trim(txt_Acc_RazSoc.Text)) = 0 Then
      MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Acc_RazSoc)
      Exit Function
   End If

   If Len(Trim(txt_Acc_NomCom.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Acc_NomCom)
      Exit Function
   End If

   If cmb_Acc_GirCom.ListIndex = -1 Then
      MsgBox "Seleccione el Giro Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Acc_GirCom)
      Exit Function
   End If

   If txt_Acc_GirCom.Enabled Then
      If Len(Trim(txt_Acc_GirCom.Text)) = 0 Then
         MsgBox "Debe ingresar el Giro Comercial.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Acc_GirCom)
         Exit Function
      End If
   End If
   
   If cmb_Acc_TipVia.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Vía de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Acc_TipVia)
      Exit Function
   End If

   If Len(Trim(txt_Acc_NomVia.Text)) = 0 Then
      MsgBox "Ingrese el Nombre de la Vóa de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Acc_NomVia)
      Exit Function
   End If
   
   If Len(Trim(txt_Acc_Numero.Text)) = 0 Then
      MsgBox "Ingrese el Número en la Vía de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Acc_Numero)
      Exit Function
   End If
   
   If cmb_Acc_TipZon.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Zona de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Acc_TipZon)
      Exit Function
   End If

   If cmb_Acc_TipZon.ItemData(cmb_Acc_TipZon.ListIndex) <> 12 Then
      If Len(Trim(txt_Acc_NomZon.Text)) = 0 Then
         MsgBox "Ingrese el Nombre de la Zona de la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Acc_NomZon)
         Exit Function
      End If
   End If

   If cmb_Acc_DptDir.ListIndex = -1 Then
      MsgBox "Seleccione el Departamento de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Acc_DptDir)
      Exit Function
   End If
   
   If cmb_Acc_PrvDir.ListIndex = -1 Then
      MsgBox "Seleccione la Provincia de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Acc_PrvDir)
      Exit Function
   End If
   
   If cmb_Acc_DstDir.ListIndex = -1 Then
      MsgBox "Seleccione el Distrito de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Acc_DstDir)
      Exit Function
   End If
   
   If Len(Trim(txt_Acc_Telef1.Text)) = 0 Then
      MsgBox "Ingrese el Teléfono de la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Acc_Telef1)
      Exit Function
   End If
   
   If CDbl(ipp_Acc_IngNet.Text) = 0 Then
      MsgBox "Ingrese el Ingreso Neto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Acc_IngNet)
      Exit Function
   End If
   
   If CDbl(ipp_Acc_PorAcc.Text) = 0 Then
      MsgBox "Ingrese el Porcentaje de Accionariado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Acc_PorAcc)
      Exit Function
   End If
   
   'Validar Fecha de Antigüedad
   
   ff_Valida_Acc = True
End Function

Private Function ff_Valida_Ind() As Integer
   ff_Valida_Ind = False
   
   If cmb_Ind_TipDoc.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Ind_TipDoc)
      Exit Function
   End If
   
   If Len(Trim(txt_Ind_NumDoc.Text)) <> 11 Then
      MsgBox "Ingrese correctamente el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ind_NumDoc)
      Exit Function
   End If

   If Not gf_Valida_RUC(Mid(txt_Ind_NumDoc.Text, 1, Len(txt_Ind_NumDoc.Text) - 1), Right(txt_Ind_NumDoc.Text, 1)) Then
      MsgBox "El Número de Documento de Identidad no es válido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ind_NumDoc)
      Exit Function
   End If
     
   If cmb_Ind_GirCom.ListIndex = -1 Then
      MsgBox "Seleccione el Giro Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Ind_GirCom)
      Exit Function
   End If

   If txt_Ind_GirCom.Enabled Then
      If Len(Trim(txt_Ind_GirCom.Text)) = 0 Then
         MsgBox "Debe ingresar el Giro Comercial.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ind_GirCom)
         Exit Function
      End If
   End If
   
   If cmb_Ind_TipVia.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Vía de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Ind_TipVia)
      Exit Function
   End If

   If Len(Trim(txt_Ind_NomVia.Text)) = 0 Then
      MsgBox "Ingrese el Nombre de la Vóa de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ind_NomVia)
      Exit Function
   End If
   
   If Len(Trim(txt_Ind_Numero.Text)) = 0 Then
      MsgBox "Ingrese el Número en la Vía de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ind_Numero)
      Exit Function
   End If
   
   If cmb_Ind_TipZon.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Zona de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Ind_TipZon)
      Exit Function
   End If

   If cmb_Ind_TipZon.ItemData(cmb_Ind_TipZon.ListIndex) <> 12 Then
      If Len(Trim(txt_Ind_NomZon.Text)) = 0 Then
         MsgBox "Ingrese el Nombre de la Zona de la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ind_NomZon)
         Exit Function
      End If
   End If

   If cmb_Ind_DptDir.ListIndex = -1 Then
      MsgBox "Seleccione el Departamento de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Ind_DptDir)
      Exit Function
   End If
   
   If cmb_Ind_PrvDir.ListIndex = -1 Then
      MsgBox "Seleccione la Provincia de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Ind_PrvDir)
      Exit Function
   End If
   
   If cmb_Ind_DstDir.ListIndex = -1 Then
      MsgBox "Seleccione el Distrito de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Ind_DstDir)
      Exit Function
   End If
   
   If Len(Trim(txt_Ind_Telef1.Text)) = 0 Then
      MsgBox "Ingrese el Teléfono de la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ind_Telef1)
      Exit Function
   End If
   
   If CDbl(ipp_Ind_IngNet.Text) = 0 Then
      MsgBox "Ingrese el Ingreso Neto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Com_IngNet)
      Exit Function
   End If
   
   'Validar Fecha de Inicio de Actividades
   
   If cmb_Ind_ConLoc.ListIndex = -1 Then
      MsgBox "Seleccione si tiene Contrato de Locación de Servicios.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Ind_ConLoc)
      Exit Function
   End If
   
   If cmb_Ind_ConLoc.ItemData(cmb_Ind_ConLoc.ListIndex) = 1 Then
      If cmb_Ind_TDoEmp.ListIndex = -1 Then
         MsgBox "Seleccione el Tipo de Documento de la Empresa donde tiene Contrato de Locación.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Ind_TDoEmp)
         Exit Function
      End If
   
      If Len(Trim(txt_Ind_NDoEmp.Text)) <> 11 Then
         MsgBox "Ingrese correctamente el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ind_NDoEmp)
         Exit Function
      End If
   
      If Not gf_Valida_RUC(Mid(txt_Ind_NDoEmp.Text, 1, Len(txt_Ind_NDoEmp.Text) - 1), Right(txt_Ind_NDoEmp.Text, 1)) Then
         MsgBox "El Número de Documento de Identidad no es válido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ind_NDoEmp)
         Exit Function
      End If
      
      If Len(Trim(pnl_Ind_FlgEmp.Caption)) = 0 Then
         MsgBox "Debe buscar la empresa en el Maestro de Empresas.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Ind_BusEmp)
         Exit Function
      End If
      
      If Len(Trim(txt_Ind_RazSoc.Text)) = 0 Then
         MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ind_RazSoc)
         Exit Function
      End If
      
      If Len(Trim(txt_Ind_Tl1Emp.Text)) = 0 Then
         MsgBox "Ingrese el Teléfono de la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ind_Tl1Emp)
         Exit Function
      End If
   
      If cmb_Ind_NomCar.ListIndex = -1 Then
         MsgBox "Seleccione el Cargo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Ind_NomCar)
         Exit Function
      End If
      
      If txt_Ind_NomCar.Enabled Then
         If Len(Trim(txt_Ind_NomCar.Text)) = 0 Then
            MsgBox "Debe ingresar el Cargo.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Ind_NomCar)
            Exit Function
         End If
      End If
      
   End If
   
   ff_Valida_Ind = True
End Function

Private Function ff_Valida_Ren() As Integer
   ff_Valida_Ren = False
   
   If cmb_Ren_TipDoc.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Ren_TipDoc)
      Exit Function
   End If
   
   If Len(Trim(txt_Ren_NumDoc.Text)) <> 11 Then
      MsgBox "Ingrese correctamente el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ren_NumDoc)
      Exit Function
   End If

   If Not gf_Valida_RUC(Mid(txt_Ren_NumDoc.Text, 1, Len(txt_Ren_NumDoc.Text) - 1), Right(txt_Ren_NumDoc.Text, 1)) Then
      MsgBox "El Número de Documento de Identidad no es válido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ren_NumDoc)
      Exit Function
   End If

   If cmb_Ren_GirCom.ListIndex = -1 Then
      MsgBox "Seleccione el Giro Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Ren_GirCom)
      Exit Function
   End If

   If Len(Trim(txt_Ren_Direc1.Text)) = 0 Then
      MsgBox "Ingrese la Dirección de la Propiedad Nro. 1.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ren_Direc1)
      Exit Function
   End If
   
   If Len(Trim(txt_Ren_NomAr1.Text)) = 0 Then
      MsgBox "Ingrese la Nombre del Arrendatario de la Propiedad Nro. 1.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ren_NomAr1)
      Exit Function
   End If
   
   If Len(Trim(txt_Ren_Tele11.Text)) = 0 Then
      MsgBox "Ingrese el Teléfono del Arrendatario de la Propiedad Nro. 1.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Ren_Tele11)
      Exit Function
   End If
   
   If CDbl(ipp_Ren_AlqMe1.Text) = 0 Then
      MsgBox "Ingrese el Importe de Alquiler Mensual de la Propiedad Nro. 1.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Ren_AlqMe1)
      Exit Function
   End If
      
   'Validar Fecha de Antigüedad de Local
   
   If chk_Alqui2.Value = 1 Then
      If Len(Trim(txt_Ren_Direc2.Text)) = 0 Then
         MsgBox "Ingrese la Dirección de la Propiedad Nro. 2.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ren_Direc2)
         Exit Function
      End If
      
      If Len(Trim(txt_Ren_NomAr2.Text)) = 0 Then
         MsgBox "Ingrese la Nombre del Arrendatario de la Propiedad Nro. 2.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ren_NomAr2)
         Exit Function
      End If
      
      If Len(Trim(txt_Ren_Tele12.Text)) = 0 Then
         MsgBox "Ingrese el Teléfono del Arrendatario de la Propiedad Nro. 2.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ren_Tele12)
         Exit Function
      End If
      
      If CDbl(ipp_Ren_AlqMe2.Text) = 0 Then
         MsgBox "Ingrese el Importe de Alquiler Mensual de la Propiedad Nro. 2.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_Ren_AlqMe2)
         Exit Function
      End If
         
      'Validar Fecha de Antigüedad de Local
   
   End If
   
   If chk_Alqui3.Value = 1 Then
      If Len(Trim(txt_Ren_Direc3.Text)) = 0 Then
         MsgBox "Ingrese la Dirección de la Propiedad Nro. 3.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ren_Direc2)
         Exit Function
      End If
      
      If Len(Trim(txt_Ren_NomAr3.Text)) = 0 Then
         MsgBox "Ingrese la Nombre del Arrendatario de la Propiedad Nro. 3.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ren_NomAr3)
         Exit Function
      End If
      
      If Len(Trim(txt_Ren_Tele13.Text)) = 0 Then
         MsgBox "Ingrese el Teléfono del Arrendatario de la Propiedad Nro. 3.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Ren_Tele13)
         Exit Function
      End If
      
      If CDbl(ipp_Ren_AlqMe3.Text) = 0 Then
         MsgBox "Ingrese el Importe de Alquiler Mensual de la Propiedad Nro. 3.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_Ren_AlqMe3)
         Exit Function
      End If
         
      'Validar Fecha de Antigüedad de Local
   
   End If
   
   If CDbl(ipp_Ren_IngNet.Text) = 0 Then
      MsgBox "Ingrese el Ingreso Neto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Ren_IngNet)
      Exit Function
   End If
   
   ff_Valida_Ren = True
End Function

Private Sub txt_Ind_NDoEmp_Change()
   pnl_Ind_FlgEmp.Caption = ""
   
   txt_Ind_RazSoc.Enabled = True
   txt_Ind_Tl1Emp.Enabled = True
   txt_Ind_Tl2Emp.Enabled = True
End Sub

Private Sub txt_Ind_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_Ind_NumDoc)
End Sub

Private Sub txt_Ind_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Ind_GirCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_Ren_GirCom_Change()
   l_str_Ren_GirCom = cmb_Ren_GirCom.Text
End Sub

Private Sub cmb_Ren_GirCom_Click()
   If cmb_Ren_GirCom.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Ren_Direc1)
      End If
   End If
End Sub

Private Sub cmb_Ren_GirCom_GotFocus()
   l_int_FlgCmb = True
   l_str_Ren_GirCom = cmb_Ren_GirCom.Text
End Sub

Private Sub cmb_Ren_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Ren_GirCom, l_str_Ren_GirCom)
      l_int_FlgCmb = True
      
      If cmb_Ren_GirCom.ListIndex > -1 Then
         l_str_Ren_GirCom = ""
      End If
      
      Call gs_SetFocus(txt_Ren_Direc1)
   End If
End Sub

Private Sub cmb_Ind_GirCom_Change()
   l_str_Ind_GirCom = cmb_Ind_GirCom.Text
End Sub

Private Sub cmb_Ind_GirCom_Click()
   txt_Ind_GirCom.Enabled = False
   txt_Ind_GirCom.Text = ""
   
   If cmb_Ind_GirCom.ListIndex > -1 Then
      If l_int_FlgCmb Then
         If l_arr_Ind_GirCom(cmb_Ind_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
            txt_Ind_GirCom.Enabled = True
            Call gs_SetFocus(txt_Ind_GirCom)
         Else
            Call gs_SetFocus(cmb_Ind_TipVia)
         End If
      End If
   End If
End Sub

Private Sub cmb_Ind_GirCom_GotFocus()
   l_int_FlgCmb = True
   l_str_Ind_GirCom = cmb_Ind_GirCom.Text
End Sub

Private Sub cmb_Ind_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ./*+#,()" + Chr(34))
   Else
      txt_Ind_GirCom.Enabled = False
      txt_Ind_GirCom.Text = ""
      
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Ind_GirCom, l_str_Ind_GirCom)
      l_int_FlgCmb = True
      
      If cmb_Ind_GirCom.ListIndex > -1 Then
         l_str_Ind_GirCom = ""
      End If
      
      If l_arr_Ind_GirCom(cmb_Ind_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
         txt_Ind_GirCom.Enabled = True
         Call gs_SetFocus(txt_Ind_GirCom)
      Else
         Call gs_SetFocus(cmb_Ind_TipVia)
      End If
   End If
End Sub

Private Sub txt_Ind_GirCom_GotFocus()
   Call gs_SelecTodo(txt_Ind_GirCom)
End Sub

Private Sub txt_Ind_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Ind_TipVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub cmb_Ind_TipVia_Click()
   Call gs_SetFocus(txt_Ind_NomVia)
End Sub

Private Sub cmb_Ind_TipVia_KeyPress(KeyAscii As Integer)
   Call cmb_Ind_TipVia_Click
End Sub

Private Sub txt_Ind_NomVia_GotFocus()
   Call gs_SelecTodo(txt_Ind_NomVia)
End Sub

Private Sub txt_Ind_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ind_Numero)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Ind_Numero_GotFocus()
   Call gs_SelecTodo(txt_Ind_Numero)
End Sub

Private Sub txt_Ind_Numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ind_Interi)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Ind_Interi_GotFocus()
   Call gs_SelecTodo(txt_Ind_Interi)
End Sub

Private Sub txt_Ind_Interi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Ind_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub cmb_Ind_TipZon_Click()
   Call gs_SetFocus(txt_Ind_NomZon)
End Sub

Private Sub cmb_Ind_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_Ind_TipZon_Click
End Sub

Private Sub txt_Ind_NomZon_GotFocus()
   Call gs_SelecTodo(txt_Ind_NomZon)
End Sub

Private Sub txt_Ind_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Ind_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Ind_Refere_GotFocus()
   Call gs_SelecTodo(txt_Ind_Refere)
End Sub

Private Sub txt_Ind_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ind_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Ind_Telef1_GotFocus()
   Call gs_SelecTodo(txt_Ind_Telef1)
End Sub

Private Sub txt_Ind_Telef1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ind_Telef2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ind_Telef2_GotFocus()
   Call gs_SelecTodo(txt_Ind_Telef2)
End Sub

Private Sub txt_Ind_Telef2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ind_NumFax)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ind_NumFax_GotFocus()
   Call gs_SelecTodo(txt_Ind_NumFax)
End Sub

Private Sub txt_Ind_NumFax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Ind_IngNet)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ind_NDoEmp_GotFocus()
   Call gs_SelecTodo(txt_Ind_NDoEmp)
End Sub

Private Sub txt_Ind_NDoEmp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Ind_BusEmp)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ind_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_Ind_RazSoc)
End Sub

Private Sub txt_Ind_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ind_Tl1Emp)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_Ind_Tl1Emp_GotFocus()
   Call gs_SelecTodo(txt_Ind_Tl1Emp)
End Sub

Private Sub txt_Ind_Tl1Emp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ind_Tl2Emp)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ind_Tl2Emp_GotFocus()
   Call gs_SelecTodo(txt_Ind_Tl2Emp)
End Sub

Private Sub txt_Ind_Tl2Emp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Ind_NomCar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_Ind_NomCar_Change()
   l_str_Ind_NomCar = cmb_Ind_NomCar.Text
End Sub

Private Sub cmb_Ind_NomCar_Click()
   txt_Ind_NomCar.Enabled = False
   txt_Ind_NomCar.Text = ""
   
   If cmb_Ind_NomCar.ListIndex > -1 Then
      If l_int_FlgCmb Then
         If l_arr_Ind_NomCar(cmb_Ind_NomCar.ListIndex + 1).Genera_Codigo = "999999" Then
            txt_Ind_NomCar.Enabled = True
            Call gs_SetFocus(txt_Ind_NomCar)
         Else
            Call gs_SetFocus(ipp_Ind_FecIng)
         End If
      End If
   End If
End Sub

Private Sub cmb_Ind_NomCar_GotFocus()
   l_int_FlgCmb = True
   l_str_Ind_NomCar = cmb_Ind_NomCar.Text
End Sub

Private Sub cmb_Ind_NomCar_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ./*+#,()" + Chr(34))
   Else
      txt_Ind_NomCar.Enabled = False
      txt_Ind_NomCar.Text = ""
      
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Ind_NomCar, l_str_Ind_NomCar)
      l_int_FlgCmb = True
      
      If cmb_Ind_NomCar.ListIndex > -1 Then
         l_str_Ind_NomCar = ""
      End If
      
      If l_arr_Ind_NomCar(cmb_Ind_NomCar.ListIndex + 1).Genera_Codigo = "999999" Then
         txt_Ind_NomCar.Enabled = True
         Call gs_SetFocus(txt_Ind_NomCar)
      Else
         Call gs_SetFocus(ipp_Ind_FecIng)
      End If
   End If
End Sub

Private Sub txt_Ind_NomCar_GotFocus()
   Call gs_SelecTodo(txt_Ind_NomCar)
End Sub

Private Sub txt_Ind_NomCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Ind_FecIng)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub cmb_Acc_TipDoc_Click()
   Call gs_SetFocus(txt_Acc_NumDoc)
End Sub

Private Sub cmb_Acc_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Acc_TipDoc_Click
   End If
End Sub

Private Sub txt_Acc_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_Acc_NumDoc)
End Sub

Private Sub txt_Acc_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Acc_BusEmp)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Acc_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_Acc_RazSoc)
End Sub

Private Sub txt_Acc_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Acc_NomCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_Acc_NomCom_GotFocus()
   Call gs_SelecTodo(txt_Acc_NomCom)
End Sub

Private Sub txt_Acc_NomCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Acc_GirCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub cmb_Acc_GirCom_Change()
   l_str_Acc_GirCom = cmb_Acc_GirCom.Text
End Sub

Private Sub cmb_Acc_GirCom_Click()
   txt_Acc_GirCom.Enabled = False
   txt_Acc_GirCom.Text = ""
   
   If cmb_Acc_GirCom.ListIndex > -1 Then
      If l_int_FlgCmb Then
         If l_arr_Acc_GirCom(cmb_Acc_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
            txt_Acc_GirCom.Enabled = True
            Call gs_SetFocus(txt_Acc_GirCom)
         Else
            Call gs_SetFocus(cmb_Acc_TipVia)
         End If
      End If
   End If
End Sub

Private Sub cmb_Acc_GirCom_GotFocus()
   l_int_FlgCmb = True
   l_str_Acc_GirCom = cmb_Acc_GirCom.Text
End Sub

Private Sub cmb_Acc_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ./*+#,()" + Chr(34))
   Else
      txt_Acc_GirCom.Enabled = False
      txt_Acc_GirCom.Text = ""
      
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Acc_GirCom, l_str_Acc_GirCom)
      l_int_FlgCmb = True
      
      If cmb_Acc_GirCom.ListIndex > -1 Then
         l_str_Acc_GirCom = ""
      End If
      
      If l_arr_Acc_GirCom(cmb_Acc_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
         txt_Acc_GirCom.Enabled = True
         Call gs_SetFocus(txt_Acc_GirCom)
      Else
         Call gs_SetFocus(cmb_Acc_TipVia)
      End If
   End If
End Sub

Private Sub txt_Acc_GirCom_GotFocus()
   Call gs_SelecTodo(txt_Acc_GirCom)
End Sub

Private Sub txt_Acc_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Acc_TipVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub cmb_Acc_TipVia_Click()
   Call gs_SetFocus(txt_Acc_NomVia)
End Sub

Private Sub cmb_Acc_TipVia_KeyPress(KeyAscii As Integer)
   Call cmb_Acc_TipVia_Click
End Sub

Private Sub txt_Acc_NomVia_GotFocus()
   Call gs_SelecTodo(txt_Acc_NomVia)
End Sub

Private Sub txt_Acc_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Acc_Numero)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Acc_Numero_GotFocus()
   Call gs_SelecTodo(txt_Acc_Numero)
End Sub

Private Sub txt_Acc_Numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Acc_Interi)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Acc_Interi_GotFocus()
   Call gs_SelecTodo(txt_Acc_Interi)
End Sub

Private Sub txt_Acc_Interi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Acc_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub cmb_Acc_TipZon_Click()
   Call gs_SetFocus(txt_Acc_NomZon)
End Sub

Private Sub cmb_Acc_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_Acc_TipZon_Click
End Sub

Private Sub txt_Acc_NomZon_GotFocus()
   Call gs_SelecTodo(txt_Acc_NomZon)
End Sub

Private Sub txt_Acc_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Acc_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Acc_Refere_GotFocus()
   Call gs_SelecTodo(txt_Acc_Refere)
End Sub

Private Sub txt_Acc_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Acc_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Acc_Telef1_GotFocus()
   Call gs_SelecTodo(txt_Acc_Telef1)
End Sub

Private Sub txt_Acc_Telef1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Acc_Telef2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Acc_Telef2_GotFocus()
   Call gs_SelecTodo(txt_Acc_Telef2)
End Sub

Private Sub txt_Acc_Telef2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Acc_NumFax)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Acc_NumFax_GotFocus()
   Call gs_SelecTodo(txt_Acc_NumFax)
End Sub

Private Sub txt_Acc_NumFax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Acc_IngNet)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_Com_TipDoc_Click()
   Call gs_SetFocus(txt_Com_NumDoc)
End Sub

Private Sub cmb_Com_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Com_TipDoc_Click
   End If
End Sub

Private Sub txt_Com_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_Com_NumDoc)
End Sub

Private Sub txt_Com_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com_RazSoc)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Com_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_Com_RazSoc)
End Sub

Private Sub txt_Com_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com_NomCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_Com_NomCom_GotFocus()
   Call gs_SelecTodo(txt_Com_NomCom)
End Sub

Private Sub txt_Com_NomCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Com_GirCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub cmb_Com_GirCom_Change()
   l_str_Com_GirCom = cmb_Com_GirCom.Text
End Sub

Private Sub cmb_Com_GirCom_Click()
   txt_Com_GirCom.Enabled = False
   txt_Com_GirCom.Text = ""
   
   If cmb_Com_GirCom.ListIndex > -1 Then
      If l_int_FlgCmb Then
         If l_arr_Com_GirCom(cmb_Com_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
            txt_Com_GirCom.Enabled = True
            Call gs_SetFocus(txt_Com_GirCom)
         Else
            Call gs_SetFocus(cmb_Com_TipVia)
         End If
      End If
   End If
End Sub

Private Sub cmb_Com_GirCom_GotFocus()
   l_int_FlgCmb = True
   l_str_Com_GirCom = cmb_Com_GirCom.Text
End Sub

Private Sub cmb_Com_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ./*+#,()" + Chr(34))
   Else
      txt_Com_GirCom.Enabled = False
      txt_Com_GirCom.Text = ""
      
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Com_GirCom, l_str_Com_GirCom)
      l_int_FlgCmb = True
      
      If cmb_Com_GirCom.ListIndex > -1 Then
         l_str_Com_GirCom = ""
      End If
      
      If l_arr_Com_GirCom(cmb_Com_GirCom.ListIndex + 1).Genera_Codigo = "999999" Then
         txt_Com_GirCom.Enabled = True
         Call gs_SetFocus(txt_Com_GirCom)
      Else
         Call gs_SetFocus(cmb_Com_TipVia)
      End If
   End If
End Sub

Private Sub txt_Com_GirCom_GotFocus()
   Call gs_SelecTodo(txt_Com_GirCom)
End Sub

Private Sub txt_Com_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Com_TipVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub cmb_Com_TipVia_Click()
   Call gs_SetFocus(txt_Com_NomVia)
End Sub

Private Sub cmb_Com_TipVia_KeyPress(KeyAscii As Integer)
   Call cmb_Com_TipVia_Click
End Sub

Private Sub txt_Com_NomVia_GotFocus()
   Call gs_SelecTodo(txt_Com_NomVia)
End Sub

Private Sub txt_Com_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com_Numero)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Com_Numero_GotFocus()
   Call gs_SelecTodo(txt_Com_Numero)
End Sub

Private Sub txt_Com_Numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com_Interi)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Com_Interi_GotFocus()
   Call gs_SelecTodo(txt_Com_Interi)
End Sub

Private Sub txt_Com_Interi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Com_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub cmb_Com_TipZon_Click()
   Call gs_SetFocus(txt_Com_NomZon)
End Sub

Private Sub cmb_Com_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_Com_TipZon_Click
End Sub

Private Sub txt_Com_NomZon_GotFocus()
   Call gs_SelecTodo(txt_Com_NomZon)
End Sub

Private Sub txt_Com_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Com_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Com_Refere_GotFocus()
   Call gs_SelecTodo(txt_Com_Refere)
End Sub

Private Sub txt_Com_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Com_Telef1_GotFocus()
   Call gs_SelecTodo(txt_Com_Telef1)
End Sub

Private Sub txt_Com_Telef1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com_Telef2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Com_Telef2_GotFocus()
   Call gs_SelecTodo(txt_Com_Telef2)
End Sub

Private Sub txt_Com_Telef2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com_NumFax)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Com_NumFax_GotFocus()
   Call gs_SelecTodo(txt_Com_NumFax)
End Sub

Private Sub txt_Com_NumFax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Com_IngNet)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub


Private Sub txt_Com_NomArr_GotFocus()
   Call gs_SelecTodo(txt_Com_NomArr)
End Sub

Private Sub txt_Com_NomArr_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com_Tl1Arr)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ().,;:/!#$º")
   End If
End Sub

Private Sub txt_Com_Tl1Arr_GotFocus()
   Call gs_SelecTodo(txt_Com_Tl1Arr)
End Sub

Private Sub txt_Com_Tl1Arr_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com_Tl2Arr)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-()")
   End If
End Sub

Private Sub txt_Com_Tl2Arr_GotFocus()
   Call gs_SelecTodo(txt_Com_Tl2Arr)
End Sub

Private Sub txt_Com_Tl2Arr_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Agrega)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_Ren_TipDoc_Click()
   Call gs_SetFocus(txt_Ren_NumDoc)
End Sub

Private Sub cmb_Ren_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Ren_TipDoc_Click
   End If
End Sub

Private Sub txt_Ren_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_Ren_NumDoc)
End Sub

Private Sub txt_Ren_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Ren_GirCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ren_Direc1_GotFocus()
   Call gs_SelecTodo(txt_Ren_Direc1)
End Sub

Private Sub txt_Ren_Direc1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ren_NomAr1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Ren_NomAr1_GotFocus()
   Call gs_SelecTodo(txt_Ren_NomAr1)
End Sub

Private Sub txt_Ren_NomAr1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ren_Tele11)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Ren_Tele11_GotFocus()
   Call gs_SelecTodo(txt_Ren_Tele11)
End Sub

Private Sub txt_Ren_Tele11_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ren_Tele21)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ren_Tele12_GotFocus()
   Call gs_SelecTodo(txt_Ren_Tele12)
End Sub

Private Sub txt_Ren_Tele12_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ren_Tele22)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ren_Direc2_GotFocus()
   Call gs_SelecTodo(txt_Ren_Direc2)
End Sub

Private Sub txt_Ren_Direc2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ren_NomAr2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Ren_NomAr2_GotFocus()
   Call gs_SelecTodo(txt_Ren_NomAr2)
End Sub

Private Sub txt_Ren_NomAr2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ren_Tele12)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Ren_Tele21_GotFocus()
   Call gs_SelecTodo(txt_Ren_Tele21)
End Sub

Private Sub txt_Ren_Tele21_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Ren_AlqMe1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ren_Tele22_GotFocus()
   Call gs_SelecTodo(txt_Ren_Tele22)
End Sub

Private Sub txt_Ren_Tele22_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Ren_AlqMe2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ren_Direc3_GotFocus()
   Call gs_SelecTodo(txt_Ren_Direc3)
End Sub

Private Sub txt_Ren_Direc3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ren_NomAr3)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Ren_NomAr3_GotFocus()
   Call gs_SelecTodo(txt_Ren_NomAr3)
End Sub

Private Sub txt_Ren_NomAr3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ren_Tele13)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Ren_Tele13_GotFocus()
   Call gs_SelecTodo(txt_Ren_Tele13)
End Sub

Private Sub txt_Ren_Tele13_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ren_Tele23)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ren_Tele23_GotFocus()
   Call gs_SelecTodo(txt_Ren_Tele23)
End Sub

Private Sub txt_Ren_Tele23_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Ren_AlqMe3)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub fs_Activa_Com(ByVal p_Activa As Integer)
   cmb_Com_TipDoc.Enabled = p_Activa
   txt_Com_NumDoc.Enabled = p_Activa
   
   cmd_Com_BusEmp.Enabled = p_Activa
   pnl_Com_FlgEmp.Enabled = p_Activa
   
   txt_Com_RazSoc.Enabled = p_Activa
   txt_Com_NomCom.Enabled = p_Activa
                 
   cmb_Com_GirCom.Enabled = p_Activa
   txt_Com_GirCom.Enabled = p_Activa
   cmb_Com_TipVia.Enabled = p_Activa
   txt_Com_NomVia.Enabled = p_Activa
   txt_Com_Numero.Enabled = p_Activa
   txt_Com_Interi.Enabled = p_Activa
   cmb_Com_TipZon.Enabled = p_Activa
   txt_Com_NomZon.Enabled = p_Activa
   cmb_Com_DptDir.Enabled = p_Activa
   cmb_Com_PrvDir.Enabled = p_Activa
   cmb_Com_DstDir.Enabled = p_Activa
   txt_Com_Refere.Enabled = p_Activa
   txt_Com_Telef1.Enabled = p_Activa
   txt_Com_Telef2.Enabled = p_Activa
   txt_Com_NumFax.Enabled = p_Activa
                 
   ipp_Com_IngNet.Enabled = p_Activa
   ipp_Com_VtaMen.Enabled = p_Activa
                 
   ipp_Com_FecIni.Enabled = p_Activa
                 
   cmb_Com_RegTri.Enabled = p_Activa
   ipp_Com_PorPar.Enabled = p_Activa
   cmb_Com_TipLoc.Enabled = p_Activa
   ipp_Com_AlqMen.Enabled = p_Activa
   txt_Com_NomArr.Enabled = p_Activa
   txt_Com_Tl1Arr.Enabled = p_Activa
   txt_Com_Tl2Arr.Enabled = p_Activa
                 
   ipp_Com_AlqMen.Enabled = p_Activa
   txt_Com_NomArr.Enabled = p_Activa
   txt_Com_Tl1Arr.Enabled = p_Activa
   txt_Com_Tl2Arr.Enabled = p_Activa
End Sub

Private Sub fs_Activa_Dep(ByVal p_Activa As Integer)
   cmb_Dep_TipDoc.Enabled = p_Activa
   txt_Dep_NumDoc.Enabled = p_Activa
   cmd_Dep_BusEmp.Enabled = p_Activa
   pnl_Dep_FlgEmp.Enabled = p_Activa
                  
   txt_Dep_RazSoc.Enabled = p_Activa
   txt_Dep_NomCom.Enabled = p_Activa
   cmb_Dep_GirCom.Enabled = p_Activa
   txt_Dep_GirCom.Enabled = p_Activa
   txt_Dep_GirCom.Enabled = p_Activa
   chk_Dep_Sucurs.Enabled = p_Activa
   txt_Dep_Sucurs.Enabled = p_Activa
   cmb_Dep_TipVia.Enabled = p_Activa
   txt_Dep_NomVia.Enabled = p_Activa
   txt_Dep_Numero.Enabled = p_Activa
   txt_Dep_Interi.Enabled = p_Activa
   cmb_Dep_TipZon.Enabled = p_Activa
   txt_Dep_NomZon.Enabled = p_Activa
   cmb_Dep_DptDir.Enabled = p_Activa
   cmb_Dep_PrvDir.Enabled = p_Activa
   cmb_Dep_DstDir.Enabled = p_Activa
   txt_Dep_Refere.Enabled = p_Activa
   txt_Dep_Telef1.Enabled = p_Activa
   txt_Dep_Telef2.Enabled = p_Activa
   txt_Dep_NumFax.Enabled = p_Activa
   txt_Dep_TeleRH.Enabled = p_Activa
   txt_Dep_AnexRH.Enabled = p_Activa
   
   ipp_Dep_IngNet.Enabled = p_Activa
   cmb_Dep_FreHab.Enabled = p_Activa
   cmb_Dep_NomCar.Enabled = p_Activa
   txt_Dep_NomCar.Enabled = p_Activa
   txt_Dep_NomCar.Enabled = p_Activa
   txt_Dep_NomAre.Enabled = p_Activa
   ipp_Dep_FecIng.Enabled = p_Activa
   txt_Dep_NumAnx.Enabled = p_Activa
   txt_Dep_TelDir.Enabled = p_Activa
   txt_Dep_Celula.Enabled = p_Activa
   txt_Dep_DirEle.Enabled = p_Activa
   ipp_Dep_FecCes.Enabled = p_Activa
End Sub

Private Sub fs_Activa_Ind(ByVal p_Activa As Integer)
   cmb_Ind_TipDoc.Enabled = p_Activa
   txt_Ind_NumDoc.Enabled = p_Activa
   cmb_Ind_GirCom.Enabled = p_Activa
   txt_Ind_GirCom.Enabled = p_Activa
   cmb_Ind_TipVia.Enabled = p_Activa
   txt_Ind_NomVia.Enabled = p_Activa
   txt_Ind_Numero.Enabled = p_Activa
   txt_Ind_Interi.Enabled = p_Activa
   cmb_Ind_TipZon.Enabled = p_Activa
   txt_Ind_NomZon.Enabled = p_Activa
   cmb_Ind_DptDir.Enabled = p_Activa
   cmb_Ind_PrvDir.Enabled = p_Activa
   cmb_Ind_DstDir.Enabled = p_Activa
   txt_Ind_Refere.Enabled = p_Activa
   txt_Ind_Telef1.Enabled = p_Activa
   txt_Ind_Telef2.Enabled = p_Activa
   txt_Ind_NumFax.Enabled = p_Activa
                  
   ipp_Ind_IngNet.Enabled = p_Activa
   cmb_Ind_ConLoc.Enabled = p_Activa
   cmb_Ind_TDoEmp.Enabled = p_Activa
   txt_Ind_NDoEmp.Enabled = p_Activa
                  
   cmd_Ind_BusEmp.Enabled = p_Activa
   cmd_Ind_BusEmp.Enabled = p_Activa
                  
   pnl_Ind_FlgEmp.Enabled = p_Activa
   cmd_Ind_BusEmp.Enabled = p_Activa
                  
   txt_Ind_RazSoc.Enabled = p_Activa
   txt_Ind_Tl1Emp.Enabled = p_Activa
   txt_Ind_Tl2Emp.Enabled = p_Activa
                  
   cmb_Ind_NomCar.Enabled = p_Activa
   txt_Ind_NomCar.Enabled = p_Activa
   ipp_Ind_FecIng.Enabled = p_Activa
End Sub

Private Sub fs_Activa_Acc(ByVal p_Activa As Integer)
   cmb_Acc_TipDoc.Enabled = p_Activa
   txt_Acc_NumDoc.Enabled = p_Activa
   
   cmd_Acc_BusEmp.Enabled = p_Activa
   pnl_Acc_FlgEmp.Enabled = p_Activa
   
   txt_Acc_RazSoc.Enabled = p_Activa
   txt_Acc_NomCom.Enabled = p_Activa
                  
   cmb_Acc_GirCom.Enabled = p_Activa
   txt_Acc_GirCom.Enabled = p_Activa
   cmb_Acc_TipVia.Enabled = p_Activa
   txt_Acc_NomVia.Enabled = p_Activa
   txt_Acc_Numero.Enabled = p_Activa
   txt_Acc_Interi.Enabled = p_Activa
   cmb_Acc_TipZon.Enabled = p_Activa
   txt_Acc_NomZon.Enabled = p_Activa
   cmb_Acc_DptDir.Enabled = p_Activa
   cmb_Acc_PrvDir.Enabled = p_Activa
   cmb_Acc_DstDir.Enabled = p_Activa
   txt_Acc_Refere.Enabled = p_Activa
   txt_Acc_Telef1.Enabled = p_Activa
   txt_Acc_Telef2.Enabled = p_Activa
   txt_Acc_NumFax.Enabled = p_Activa
                  
   ipp_Acc_IngNet.Enabled = p_Activa
   ipp_Acc_PorAcc.Enabled = p_Activa
   ipp_Acc_FecAnt.Enabled = p_Activa
End Sub

Private Sub fs_Activa_Ren(ByVal p_Activa As Integer)
   cmb_Ren_TipDoc.Enabled = p_Activa
   txt_Ren_NumDoc.Enabled = p_Activa
   cmb_Ren_GirCom.Enabled = p_Activa
                  
   txt_Ren_Direc1.Enabled = p_Activa
   txt_Ren_NomAr1.Enabled = p_Activa
   txt_Ren_Tele11.Enabled = p_Activa
   txt_Ren_Tele21.Enabled = p_Activa
   ipp_Ren_AlqMe1.Enabled = p_Activa
   ipp_Ren_FIAlq1.Enabled = p_Activa
   
   chk_Alqui2.Enabled = p_Activa
   txt_Ren_Direc2.Enabled = p_Activa
   txt_Ren_NomAr2.Enabled = p_Activa
   txt_Ren_Tele12.Enabled = p_Activa
   txt_Ren_Tele22.Enabled = p_Activa
   ipp_Ren_AlqMe2.Enabled = p_Activa
   ipp_Ren_FIAlq2.Enabled = p_Activa
                  
   txt_Ren_Direc2.Enabled = p_Activa
   txt_Ren_NomAr2.Enabled = p_Activa
   txt_Ren_Tele12.Enabled = p_Activa
   txt_Ren_Tele22.Enabled = p_Activa
   ipp_Ren_AlqMe2.Enabled = p_Activa
   ipp_Ren_FIAlq2.Enabled = p_Activa
   
   chk_Alqui3.Enabled = p_Activa
   txt_Ren_Direc3.Enabled = p_Activa
   txt_Ren_NomAr3.Enabled = p_Activa
   txt_Ren_Tele13.Enabled = p_Activa
   txt_Ren_Tele23.Enabled = p_Activa
   ipp_Ren_AlqMe3.Enabled = p_Activa
   ipp_Ren_FIAlq3.Enabled = p_Activa
                  
   txt_Ren_Direc3.Enabled = p_Activa
   txt_Ren_NomAr3.Enabled = p_Activa
   txt_Ren_Tele13.Enabled = p_Activa
   txt_Ren_Tele23.Enabled = p_Activa
   ipp_Ren_AlqMe3.Enabled = p_Activa
   ipp_Ren_FIAlq3.Enabled = p_Activa
                  
   ipp_Ren_IngNet.Enabled = p_Activa
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_ActEco.Enabled = p_Activa
   cmb_OrdAct.Enabled = p_Activa

   cmd_Agrega.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_NueAct.Enabled = Not p_Activa
   cmd_BorAct.Enabled = Not p_Activa
   cmd_EdiAct.Enabled = Not p_Activa
End Sub

Private Sub fs_Limpia()
   cmb_ActEco.ListIndex = -1
   cmb_OrdAct.ListIndex = -1
   
   Call fs_Limpia_Dep
   Call fs_Limpia_Ind
   Call fs_Limpia_Com
   Call fs_Limpia_Acc
   Call fs_Limpia_Ren
End Sub

Private Sub fs_Carga_Arreglo_Tit()
   Dim r_int_Contad     As Integer
   
   For r_int_Contad = 1 To UBound(modatecli_g_arr_Tit_ActEco)
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      'Obteniendo Descripción de Orden de Actividad
      grd_Listad.Col = 0
      grd_Listad.Text = moddat_gf_Consulta_Pardes("007", modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_OrdAct)
      
      'Obteniendo Descripción de Código de Actividad
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_Pardes("008", modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_CodAct)
   
      grd_Listad.Col = 2
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_OrdAct

      grd_Listad.Col = 3
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_CodAct
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TipDoc)
   
      grd_Listad.Col = 5
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NumDoc
   
      grd_Listad.Col = 6
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_RazSoc
   
      grd_Listad.Col = 7
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NomCom
   
      grd_Listad.Col = 8
      grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_CodCiu)
      
      grd_Listad.Col = 9
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_GiroCd
      
      grd_Listad.Col = 10
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_GiroNm
      
      grd_Listad.Col = 11
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Sucurs
      
      grd_Listad.Col = 12
      grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TipVia)
      
      grd_Listad.Col = 13
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NomVia
      
      grd_Listad.Col = 14
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Numero
      
      grd_Listad.Col = 15
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Interi
      
      grd_Listad.Col = 16
      grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TipZon)
      
      grd_Listad.Col = 17
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NomZon
   
      grd_Listad.Col = 18
      grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_DptDir)
   
      grd_Listad.Col = 19
      grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_PrvDir)
   
      grd_Listad.Col = 20
      grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_DstDir)
   
      grd_Listad.Col = 21
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Refere
   
      grd_Listad.Col = 22
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Telef1
   
      grd_Listad.Col = 23
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Telef2
      
      grd_Listad.Col = 24
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NumFax
      
      grd_Listad.Col = 25
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TeleRH
   
      grd_Listad.Col = 26
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_AnexRH
   
      grd_Listad.Col = 27
      grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_SecEco
   
      Select Case modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_CodAct
         Case 11
            grd_Listad.Col = 29
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_FlgEmp
   
            grd_Listad.Col = 30
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_IngNet)
   
            grd_Listad.Col = 31
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_FreHab)
   
            grd_Listad.Col = 32
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_CargoC
   
            grd_Listad.Col = 33
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_CargoN
   
            grd_Listad.Col = 34
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_NomAre
   
            grd_Listad.Col = 35
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_FecIng
   
            grd_Listad.Col = 36
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_NumAnx
   
            grd_Listad.Col = 37
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_TelDir
            
            grd_Listad.Col = 38
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_Celula
            
            grd_Listad.Col = 39
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_DirEle
            
            If modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_OrdAct = 9 Then
               grd_Listad.Col = 40
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_FecCes
            End If
            
         Case 21
            grd_Listad.Col = 43
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_IngNet)
            
            grd_Listad.Col = 44
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_FecIni
         
            grd_Listad.Col = 45
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_ConLoc)
            
            If modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_ConLoc = 1 Then
               grd_Listad.Col = 46
               grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_TDoEmp)
            
               grd_Listad.Col = 47
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_NDoEmp
            
               grd_Listad.Col = 48
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_RazSoc
            
               grd_Listad.Col = 49
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_Tl1Emp
            
               grd_Listad.Col = 50
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_Tl2Emp
            
               grd_Listad.Col = 51
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_CargoC
            
               grd_Listad.Col = 52
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_CargoN
            
               grd_Listad.Col = 53
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_FecIng
            
               grd_Listad.Col = 54
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_FlgEmp
            End If
         
         Case 31
            grd_Listad.Col = 85
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_IngNet)
         
            grd_Listad.Col = 86
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_VtaMen)
      
            grd_Listad.Col = 87
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_FecIni
      
            grd_Listad.Col = 88
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_RegTri)
      
            grd_Listad.Col = 89
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_PorPar)
      
            grd_Listad.Col = 90
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_TipLoc)
            
            If modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_TipLoc = 2 Then
               grd_Listad.Col = 91
               grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_AlqMen)
            
               grd_Listad.Col = 92
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_NomArr
            
               grd_Listad.Col = 93
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_Tl1Arr
            
               grd_Listad.Col = 94
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_Tl2Arr
            End If
            
            grd_Listad.Col = 95
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_FlgEmp
            
      
         Case 41
            grd_Listad.Col = 57
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Acc_IngNet)
         
            grd_Listad.Col = 58
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Acc_PorAcc)
         
            grd_Listad.Col = 59
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Acc_FecAnt)
         
            grd_Listad.Col = 60
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Acc_FlgEmp
         
         Case 51
            grd_Listad.Col = 62
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Direc1
         
            grd_Listad.Col = 63
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_NomAr1
         
            grd_Listad.Col = 64
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele11
         
            grd_Listad.Col = 65
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele21
         
            grd_Listad.Col = 66
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_AlqMe1)
         
            grd_Listad.Col = 67
            grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_FIAlq1
            
            grd_Listad.Col = 68
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Alqui2)
            
            If modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Alqui2 = 1 Then
               grd_Listad.Col = 69
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Direc2
            
               grd_Listad.Col = 70
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_NomAr2
            
               grd_Listad.Col = 71
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele12
            
               grd_Listad.Col = 72
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele22
            
               grd_Listad.Col = 73
               grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_AlqMe2)
            
               grd_Listad.Col = 74
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_FIAlq2
            End If
            
            grd_Listad.Col = 75
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Alqui3)
            
            If modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Alqui3 = 1 Then
               grd_Listad.Col = 76
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Direc3
            
               grd_Listad.Col = 77
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_NomAr3
            
               grd_Listad.Col = 78
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele13
            
               grd_Listad.Col = 79
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele23
            
               grd_Listad.Col = 80
               grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_AlqMe3)
            
               grd_Listad.Col = 81
               grd_Listad.Text = modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_FIAlq3
            End If
            
            grd_Listad.Col = 82
            grd_Listad.Text = CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_IngNet)
      End Select
   Next r_int_Contad
End Sub

Private Sub fs_Carga_Arreglo_Cyg()
   Dim r_int_Contad     As Integer
   
   For r_int_Contad = 1 To UBound(modatecli_g_arr_Cyg_ActEco)
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      
      'Obteniendo Descripción de Orden de Actividad
      grd_Listad.Col = 0
      grd_Listad.Text = moddat_gf_Consulta_Pardes("007", modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_OrdAct)
      
      'Obteniendo Descripción de Código de Actividad
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_Pardes("008", modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_CodAct)
      
      grd_Listad.Col = 2
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_OrdAct
   
      grd_Listad.Col = 3
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_CodAct
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TipDoc)
   
      grd_Listad.Col = 5
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NumDoc
   
      grd_Listad.Col = 6
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_RazSoc
   
      grd_Listad.Col = 7
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NomCom
   
      grd_Listad.Col = 8
      grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_CodCiu)
      
      grd_Listad.Col = 9
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_GiroCd
      
      grd_Listad.Col = 10
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_GiroNm
      
      grd_Listad.Col = 11
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Sucurs
      
      grd_Listad.Col = 12
      grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TipVia)
      
      grd_Listad.Col = 13
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NomVia
      
      grd_Listad.Col = 14
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Numero
      
      grd_Listad.Col = 15
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Interi
      
      grd_Listad.Col = 16
      grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TipZon)
      
      grd_Listad.Col = 17
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NomZon
   
      grd_Listad.Col = 18
      grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_DptDir)
   
      grd_Listad.Col = 19
      grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_PrvDir)
   
      grd_Listad.Col = 20
      grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_DstDir)
   
      grd_Listad.Col = 21
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Refere
   
      grd_Listad.Col = 22
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Telef1
   
      grd_Listad.Col = 23
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Telef2
      
      grd_Listad.Col = 24
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NumFax
      
      grd_Listad.Col = 25
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TeleRH
   
      grd_Listad.Col = 26
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_AnexRH
   
      grd_Listad.Col = 27
      grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_SecEco
   
      Select Case modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_CodAct
         Case 11
            grd_Listad.Col = 29
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_FlgEmp
   
            grd_Listad.Col = 30
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_IngNet)
   
            grd_Listad.Col = 31
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_FreHab)
   
            grd_Listad.Col = 32
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_CargoC
   
            grd_Listad.Col = 33
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_CargoN
   
            grd_Listad.Col = 34
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_NomAre
   
            grd_Listad.Col = 35
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_FecIng
   
            grd_Listad.Col = 36
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_NumAnx
   
            grd_Listad.Col = 37
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_TelDir
            
            grd_Listad.Col = 38
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_Celula
            
            grd_Listad.Col = 39
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_DirEle
            
            If modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_OrdAct = 9 Then
               grd_Listad.Col = 40
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_FecCes
            End If
            
         Case 21
            grd_Listad.Col = 43
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_IngNet)
            
            grd_Listad.Col = 44
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_FecIni
         
            grd_Listad.Col = 45
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_ConLoc)
            
            If modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_ConLoc = 1 Then
               grd_Listad.Col = 46
               grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_TDoEmp)
            
               grd_Listad.Col = 47
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_NDoEmp
            
               grd_Listad.Col = 48
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_RazSoc
            
               grd_Listad.Col = 49
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_Tl1Emp
            
               grd_Listad.Col = 50
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_Tl2Emp
            
               grd_Listad.Col = 51
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_CargoC
            
               grd_Listad.Col = 52
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_CargoN
            
               grd_Listad.Col = 53
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_FecIng
            
               grd_Listad.Col = 54
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_FlgEmp
            End If
         
         Case 31
            grd_Listad.Col = 85
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_IngNet)
         
            grd_Listad.Col = 86
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_VtaMen)
      
            grd_Listad.Col = 87
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_FecIni
      
            grd_Listad.Col = 88
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_RegTri)
      
            grd_Listad.Col = 89
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_PorPar)
      
            grd_Listad.Col = 90
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_TipLoc)
            
            If modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_TipLoc = 2 Then
               grd_Listad.Col = 91
               grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_AlqMen)
            
               grd_Listad.Col = 92
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_NomArr
            
               grd_Listad.Col = 93
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_Tl1Arr
            
               grd_Listad.Col = 94
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_Tl2Arr
            End If
      
            grd_Listad.Col = 95
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_FlgEmp
      
         Case 41
            grd_Listad.Col = 57
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Acc_IngNet)
         
            grd_Listad.Col = 58
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Acc_PorAcc)
         
            grd_Listad.Col = 59
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Acc_FecAnt)
         
            grd_Listad.Col = 60
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Acc_FlgEmp
         
         Case 51
            grd_Listad.Col = 62
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Direc1
         
            grd_Listad.Col = 63
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_NomAr1
         
            grd_Listad.Col = 64
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele11
         
            grd_Listad.Col = 65
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele21
         
            grd_Listad.Col = 66
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_AlqMe1)
         
            grd_Listad.Col = 67
            grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_FIAlq1
            
            grd_Listad.Col = 68
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Alqui2)
            
            If modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Alqui2 = 1 Then
               grd_Listad.Col = 69
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Direc2
            
               grd_Listad.Col = 70
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_NomAr2
            
               grd_Listad.Col = 71
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele12
            
               grd_Listad.Col = 72
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele22
            
               grd_Listad.Col = 73
               grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_AlqMe2)
            
               grd_Listad.Col = 74
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_FIAlq2
            End If
            
            grd_Listad.Col = 75
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Alqui3)
            
            If modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Alqui3 = 1 Then
               grd_Listad.Col = 76
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Direc3
            
               grd_Listad.Col = 77
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_NomAr3
            
               grd_Listad.Col = 78
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele13
            
               grd_Listad.Col = 79
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele23
            
               grd_Listad.Col = 80
               grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_AlqMe3)
            
               grd_Listad.Col = 81
               grd_Listad.Text = modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_FIAlq3
            End If
            
            grd_Listad.Col = 82
            grd_Listad.Text = CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_IngNet)
      End Select
   Next r_int_Contad
End Sub

Private Sub fs_Graba_Arreglo_Tit()
   Dim r_int_Contad     As Integer
   
   ReDim modatecli_g_arr_Tit_ActEco(grd_Listad.Rows)
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
   
      grd_Listad.Col = 2
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_OrdAct = CInt(grd_Listad.Text)
      
      grd_Listad.Col = 3
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_CodAct = CInt(grd_Listad.Text)
   
      grd_Listad.Col = 4
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_TipDoc = CInt(grd_Listad.Text)
            
      grd_Listad.Col = 5
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_NumDoc = grd_Listad.Text
            
      grd_Listad.Col = 6
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_RazSoc = grd_Listad.Text
            
      grd_Listad.Col = 7
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_NomCom = grd_Listad.Text
            
      grd_Listad.Col = 8
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_CodCiu = CInt(grd_Listad.Text)
      
      grd_Listad.Col = 11
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Sucurs = grd_Listad.Text
      
      grd_Listad.Col = 9
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_GiroCd = grd_Listad.Text
      
      grd_Listad.Col = 10
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_GiroNm = grd_Listad.Text
      
      grd_Listad.Col = 12
      If Len(Trim(grd_Listad.Text)) > 0 Then
         modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_TipVia = CInt(grd_Listad.Text)
      End If
   
      grd_Listad.Col = 13
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_NomVia = grd_Listad.Text
   
      grd_Listad.Col = 14
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Numero = grd_Listad.Text
   
      grd_Listad.Col = 15
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Interi = grd_Listad.Text
   
      grd_Listad.Col = 16
      If Len(Trim(grd_Listad.Text)) > 0 Then
         modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_TipZon = CInt(grd_Listad.Text)
      End If
   
      grd_Listad.Col = 17
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_NomZon = grd_Listad.Text
   
      grd_Listad.Col = 18
      If Len(Trim(grd_Listad.Text)) > 0 Then
         modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_DptDir = CInt(grd_Listad.Text)
      End If
   
      grd_Listad.Col = 19
      If Len(Trim(grd_Listad.Text)) > 0 Then
         modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_PrvDir = CInt(grd_Listad.Text)
      End If
      
      grd_Listad.Col = 20
      If Len(Trim(grd_Listad.Text)) > 0 Then
         modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_DstDir = CInt(grd_Listad.Text)
      End If
      
      grd_Listad.Col = 21
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Refere = grd_Listad.Text
      
      grd_Listad.Col = 22
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Telef1 = grd_Listad.Text
      
      grd_Listad.Col = 23
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Telef2 = grd_Listad.Text
      
      grd_Listad.Col = 24
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_NumFax = grd_Listad.Text
      
      grd_Listad.Col = 25
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_TeleRH = grd_Listad.Text
      
      grd_Listad.Col = 26
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_AnexRH = grd_Listad.Text
      
      grd_Listad.Col = 27
      modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_SecEco = grd_Listad.Text
      
      Select Case modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_CodAct
         Case 11
            grd_Listad.Col = 29
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_FlgEmp = grd_Listad.Text
            
            grd_Listad.Col = 30
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_IngNet = CDbl(grd_Listad.Text)
            
            grd_Listad.Col = 31
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_FreHab = CInt(grd_Listad.Text)
            
            grd_Listad.Col = 32
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_CargoC = grd_Listad.Text
            
            grd_Listad.Col = 33
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_CargoN = grd_Listad.Text
            
            grd_Listad.Col = 34
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_NomAre = grd_Listad.Text
            
            grd_Listad.Col = 35
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_FecIng = grd_Listad.Text
            
            grd_Listad.Col = 36
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_NumAnx = grd_Listad.Text
            
            grd_Listad.Col = 37
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_TelDir = grd_Listad.Text
            
            grd_Listad.Col = 38
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_Celula = grd_Listad.Text
            
            grd_Listad.Col = 39
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_DirEle = grd_Listad.Text
            
            grd_Listad.Col = 40
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Dep_FecCes = grd_Listad.Text
         
         Case 21
            grd_Listad.Col = 43
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_IngNet = CDbl(grd_Listad.Text)
            
            grd_Listad.Col = 44
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_FecIni = grd_Listad.Text
            
            grd_Listad.Col = 45
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_ConLoc = CInt(grd_Listad.Text)
            
            If modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_ConLoc = 1 Then
               grd_Listad.Col = 46
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_TDoEmp = CInt(grd_Listad.Text)
            
               grd_Listad.Col = 47
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_NDoEmp = grd_Listad.Text
            
               grd_Listad.Col = 48
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_RazSoc = grd_Listad.Text
            
               grd_Listad.Col = 49
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_Tl1Emp = grd_Listad.Text
            
               grd_Listad.Col = 50
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_Tl2Emp = grd_Listad.Text
            
               grd_Listad.Col = 51
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_CargoC = grd_Listad.Text
            
               grd_Listad.Col = 52
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_CargoN = grd_Listad.Text
            
               grd_Listad.Col = 53
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_FecIng = grd_Listad.Text
            
               grd_Listad.Col = 54
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ind_FlgEmp = grd_Listad.Text
            End If
            
         Case 31
            grd_Listad.Col = 85
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_IngNet = CDbl(grd_Listad.Text)
         
            grd_Listad.Col = 86
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_VtaMen = CDbl(grd_Listad.Text)
         
            grd_Listad.Col = 87
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_FecIni = grd_Listad.Text
         
            grd_Listad.Col = 88
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_RegTri = CInt(grd_Listad.Text)
         
            grd_Listad.Col = 89
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_PorPar = CDbl(grd_Listad.Text)
         
            grd_Listad.Col = 90
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_TipLoc = CInt(grd_Listad.Text)
            
            If modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_TipLoc = 2 Then
               grd_Listad.Col = 91
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_AlqMen = CDbl(grd_Listad.Text)
            
               grd_Listad.Col = 92
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_NomArr = grd_Listad.Text
            
               grd_Listad.Col = 93
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_Tl1Arr = grd_Listad.Text
            
               grd_Listad.Col = 94
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_Tl2Arr = grd_Listad.Text
            End If
         
            grd_Listad.Col = 95
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Com_FlgEmp = grd_Listad.Text
         
         Case 41
            grd_Listad.Col = 57
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Acc_IngNet = CDbl(grd_Listad.Text)
         
            grd_Listad.Col = 58
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Acc_PorAcc = CDbl(grd_Listad.Text)
         
            grd_Listad.Col = 59
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Acc_FecAnt = grd_Listad.Text
         
            grd_Listad.Col = 60
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Acc_FlgEmp = grd_Listad.Text
         
         Case 51
            grd_Listad.Col = 62
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Direc1 = grd_Listad.Text
            
            grd_Listad.Col = 63
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_NomAr1 = grd_Listad.Text
            
            grd_Listad.Col = 64
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Tele11 = grd_Listad.Text
            
            grd_Listad.Col = 65
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Tele21 = grd_Listad.Text
            
            grd_Listad.Col = 66
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_AlqMe1 = CDbl(grd_Listad.Text)
            
            grd_Listad.Col = 67
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_FIAlq1 = grd_Listad.Text
            
            grd_Listad.Col = 68
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Alqui2 = CInt(grd_Listad.Text)
            
            If modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Alqui2 = 1 Then
               grd_Listad.Col = 69
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Direc2 = grd_Listad.Text
               
               grd_Listad.Col = 70
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_NomAr2 = grd_Listad.Text
               
               grd_Listad.Col = 71
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Tele12 = grd_Listad.Text
               
               grd_Listad.Col = 72
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Tele22 = grd_Listad.Text
               
               grd_Listad.Col = 73
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_AlqMe2 = CDbl(grd_Listad.Text)
               
               grd_Listad.Col = 74
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_FIAlq2 = grd_Listad.Text
            End If
            
            grd_Listad.Col = 75
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Alqui3 = CInt(grd_Listad.Text)
            
            If modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Alqui3 = 1 Then
               grd_Listad.Col = 76
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Direc3 = grd_Listad.Text
               
               grd_Listad.Col = 77
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_NomAr3 = grd_Listad.Text
               
               grd_Listad.Col = 78
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Tele13 = grd_Listad.Text
               
               grd_Listad.Col = 79
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_Tele23 = grd_Listad.Text
               
               grd_Listad.Col = 80
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_AlqMe3 = CDbl(grd_Listad.Text)
               
               grd_Listad.Col = 81
               modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_FIAlq3 = grd_Listad.Text
            End If
            
            grd_Listad.Col = 82
            modatecli_g_arr_Tit_ActEco(r_int_Contad + 1).ActEco_Ren_IngNet = CDbl(grd_Listad.Text)
      End Select
   Next r_int_Contad
   
   modatecli_g_int_ActEcoTit = 2

End Sub

Private Sub fs_Graba_Arreglo_Cyg()
   Dim r_int_Contad     As Integer
   
   ReDim modatecli_g_arr_Cyg_ActEco(grd_Listad.Rows)
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
   
      grd_Listad.Col = 2
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_OrdAct = CInt(grd_Listad.Text)
      
      grd_Listad.Col = 3
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_CodAct = CInt(grd_Listad.Text)
   
      grd_Listad.Col = 4
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_TipDoc = CInt(grd_Listad.Text)
            
      grd_Listad.Col = 5
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_NumDoc = grd_Listad.Text
            
      grd_Listad.Col = 6
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_RazSoc = grd_Listad.Text
            
      grd_Listad.Col = 7
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_NomCom = grd_Listad.Text
            
      grd_Listad.Col = 8
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_CodCiu = CInt(grd_Listad.Text)
      
      grd_Listad.Col = 11
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Sucurs = grd_Listad.Text
      
      grd_Listad.Col = 9
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_GiroCd = grd_Listad.Text
      
      grd_Listad.Col = 10
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_GiroNm = grd_Listad.Text
      
      grd_Listad.Col = 12
      If Len(Trim(grd_Listad.Text)) > 0 Then
         modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_TipVia = CInt(grd_Listad.Text)
      End If
   
      grd_Listad.Col = 13
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_NomVia = grd_Listad.Text
   
      grd_Listad.Col = 14
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Numero = grd_Listad.Text
   
      grd_Listad.Col = 15
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Interi = grd_Listad.Text
   
      grd_Listad.Col = 16
      If Len(Trim(grd_Listad.Text)) > 0 Then
         modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_TipZon = CInt(grd_Listad.Text)
      End If
   
      grd_Listad.Col = 17
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_NomZon = grd_Listad.Text
   
      grd_Listad.Col = 18
      If Len(Trim(grd_Listad.Text)) > 0 Then
         modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_DptDir = CInt(grd_Listad.Text)
      End If
   
      grd_Listad.Col = 19
      If Len(Trim(grd_Listad.Text)) > 0 Then
         modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_PrvDir = CInt(grd_Listad.Text)
      End If
      
      grd_Listad.Col = 20
      If Len(Trim(grd_Listad.Text)) > 0 Then
         modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_DstDir = CInt(grd_Listad.Text)
      End If
      
      grd_Listad.Col = 21
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Refere = grd_Listad.Text
      
      grd_Listad.Col = 22
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Telef1 = grd_Listad.Text
      
      grd_Listad.Col = 23
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Telef2 = grd_Listad.Text
      
      grd_Listad.Col = 24
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_NumFax = grd_Listad.Text
      
      grd_Listad.Col = 25
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_TeleRH = grd_Listad.Text
      
      grd_Listad.Col = 26
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_AnexRH = grd_Listad.Text
      
      grd_Listad.Col = 27
      modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_SecEco = grd_Listad.Text
      
      Select Case modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_CodAct
         Case 11
            grd_Listad.Col = 29
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_FlgEmp = grd_Listad.Text
            
            grd_Listad.Col = 30
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_IngNet = CDbl(grd_Listad.Text)
            
            grd_Listad.Col = 31
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_FreHab = CInt(grd_Listad.Text)
            
            grd_Listad.Col = 32
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_CargoC = grd_Listad.Text
            
            grd_Listad.Col = 33
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_CargoN = grd_Listad.Text
            
            grd_Listad.Col = 34
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_NomAre = grd_Listad.Text
            
            grd_Listad.Col = 35
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_FecIng = grd_Listad.Text
            
            grd_Listad.Col = 36
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_NumAnx = grd_Listad.Text
            
            grd_Listad.Col = 37
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_TelDir = grd_Listad.Text
            
            grd_Listad.Col = 38
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_Celula = grd_Listad.Text
            
            grd_Listad.Col = 39
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_DirEle = grd_Listad.Text
            
            grd_Listad.Col = 40
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Dep_FecCes = grd_Listad.Text
         
         Case 21
            grd_Listad.Col = 43
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_IngNet = CDbl(grd_Listad.Text)
            
            grd_Listad.Col = 44
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_FecIni = grd_Listad.Text
            
            grd_Listad.Col = 45
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_ConLoc = CInt(grd_Listad.Text)
            
            If modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_ConLoc = 1 Then
               grd_Listad.Col = 46
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_TDoEmp = CInt(grd_Listad.Text)
            
               grd_Listad.Col = 47
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_NDoEmp = grd_Listad.Text
            
               grd_Listad.Col = 48
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_RazSoc = grd_Listad.Text
            
               grd_Listad.Col = 49
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_Tl1Emp = grd_Listad.Text
            
               grd_Listad.Col = 50
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_Tl2Emp = grd_Listad.Text
            
               grd_Listad.Col = 51
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_CargoC = grd_Listad.Text
            
               grd_Listad.Col = 52
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_CargoN = grd_Listad.Text
            
               grd_Listad.Col = 53
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_FecIng = grd_Listad.Text
            
               grd_Listad.Col = 54
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ind_FlgEmp = grd_Listad.Text
            End If
            
         Case 31
            grd_Listad.Col = 85
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_IngNet = CDbl(grd_Listad.Text)
         
            grd_Listad.Col = 86
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_VtaMen = CDbl(grd_Listad.Text)
         
            grd_Listad.Col = 87
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_FecIni = grd_Listad.Text
         
            grd_Listad.Col = 88
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_RegTri = CInt(grd_Listad.Text)
         
            grd_Listad.Col = 89
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_PorPar = CDbl(grd_Listad.Text)
         
            grd_Listad.Col = 90
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_TipLoc = CInt(grd_Listad.Text)
            
            If modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_TipLoc = 2 Then
               grd_Listad.Col = 91
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_AlqMen = CDbl(grd_Listad.Text)
            
               grd_Listad.Col = 92
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_NomArr = grd_Listad.Text
            
               grd_Listad.Col = 93
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_Tl1Arr = grd_Listad.Text
            
               grd_Listad.Col = 94
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_Tl2Arr = grd_Listad.Text
            End If
         
            grd_Listad.Col = 95
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Com_FlgEmp = grd_Listad.Text
         
         Case 41
            grd_Listad.Col = 57
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Acc_IngNet = CDbl(grd_Listad.Text)
         
            grd_Listad.Col = 58
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Acc_PorAcc = CDbl(grd_Listad.Text)
         
            grd_Listad.Col = 59
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Acc_FecAnt = grd_Listad.Text
         
            grd_Listad.Col = 60
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Acc_FlgEmp = grd_Listad.Text
         
         Case 51
            grd_Listad.Col = 62
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Direc1 = grd_Listad.Text
            
            grd_Listad.Col = 63
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_NomAr1 = grd_Listad.Text
            
            grd_Listad.Col = 64
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Tele11 = grd_Listad.Text
            
            grd_Listad.Col = 65
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Tele21 = grd_Listad.Text
            
            grd_Listad.Col = 66
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_AlqMe1 = CDbl(grd_Listad.Text)
            
            grd_Listad.Col = 67
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_FIAlq1 = grd_Listad.Text
            
            grd_Listad.Col = 68
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Alqui2 = CInt(grd_Listad.Text)
            
            If modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Alqui2 = 1 Then
               grd_Listad.Col = 69
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Direc2 = grd_Listad.Text
               
               grd_Listad.Col = 70
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_NomAr2 = grd_Listad.Text
               
               grd_Listad.Col = 71
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Tele12 = grd_Listad.Text
               
               grd_Listad.Col = 72
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Tele22 = grd_Listad.Text
               
               grd_Listad.Col = 73
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_AlqMe2 = CDbl(grd_Listad.Text)
               
               grd_Listad.Col = 74
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_FIAlq2 = grd_Listad.Text
            End If
            
            grd_Listad.Col = 75
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Alqui3 = CInt(grd_Listad.Text)
            
            If modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Alqui3 = 1 Then
               grd_Listad.Col = 76
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Direc3 = grd_Listad.Text
               
               grd_Listad.Col = 77
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_NomAr3 = grd_Listad.Text
               
               grd_Listad.Col = 78
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Tele13 = grd_Listad.Text
               
               grd_Listad.Col = 79
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_Tele23 = grd_Listad.Text
               
               grd_Listad.Col = 80
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_AlqMe3 = CDbl(grd_Listad.Text)
               
               grd_Listad.Col = 81
               modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_FIAlq3 = grd_Listad.Text
            End If
            
            grd_Listad.Col = 82
            modatecli_g_arr_Cyg_ActEco(r_int_Contad + 1).ActEco_Ren_IngNet = CDbl(grd_Listad.Text)
      End Select
   Next r_int_Contad
   
   modatecli_g_int_ActEcoCyg = 2
End Sub

Private Sub fs_BusEmp(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Select Case cmb_ActEco.ItemData(cmb_ActEco.ListIndex)
         Case 11
            pnl_Dep_FlgEmp.Caption = "NR"
            
            txt_Dep_RazSoc.Enabled = True
            txt_Dep_NomCom.Enabled = True
            cmb_Dep_GirCom.Enabled = True
            txt_Dep_GirCom.Enabled = True
            cmb_Dep_TipVia.Enabled = True
            txt_Dep_NomVia.Enabled = True
            txt_Dep_Numero.Enabled = True
            txt_Dep_Interi.Enabled = True
            cmb_Dep_TipZon.Enabled = True
            txt_Dep_NomZon.Enabled = True
            cmb_Dep_DptDir.Enabled = True
            cmb_Dep_PrvDir.Enabled = True
            cmb_Dep_DstDir.Enabled = True
            txt_Dep_Refere.Enabled = True
            txt_Dep_Telef1.Enabled = True
            txt_Dep_Telef2.Enabled = True
            txt_Dep_NumFax.Enabled = True
            txt_Dep_TeleRH.Enabled = True
            txt_Dep_AnexRH.Enabled = True
            
            txt_Dep_RazSoc.Text = ""
            txt_Dep_NomCom.Text = ""
            cmb_Dep_GirCom.ListIndex = -1
            txt_Dep_GirCom.Text = ""
            txt_Dep_GirCom.Enabled = False
            chk_Dep_Sucurs.Value = 0
            txt_Dep_Sucurs.Text = ""
            txt_Dep_Sucurs.Enabled = False
            cmb_Dep_TipVia.ListIndex = -1
            txt_Dep_NomVia.Text = ""
            txt_Dep_Numero.Text = ""
            txt_Dep_Interi.Text = ""
            cmb_Dep_TipZon.ListIndex = -1
            txt_Dep_NomZon.Text = ""
            cmb_Dep_DptDir.ListIndex = -1
            cmb_Dep_PrvDir.Clear
            cmb_Dep_DstDir.Clear
            txt_Dep_Refere.Text = ""
            txt_Dep_Telef1.Text = ""
            txt_Dep_Telef2.Text = ""
            txt_Dep_NumFax.Text = ""
            txt_Dep_TeleRH.Text = ""
            txt_Dep_AnexRH.Text = ""
            
            Call gs_SetFocus(txt_Dep_RazSoc)
            
         Case 21
            pnl_Ind_FlgEmp.Caption = "NR"
            
            txt_Ind_RazSoc.Enabled = True
            txt_Ind_Tl1Emp.Enabled = True
            txt_Ind_Tl2Emp.Enabled = True
            
            txt_Ind_RazSoc.Text = ""
            txt_Ind_Tl1Emp.Text = ""
            txt_Ind_Tl2Emp.Text = ""
            
            Call gs_SetFocus(txt_Ind_RazSoc)
            
         Case 31
            pnl_Com_FlgEmp.Caption = "NR"
            
            txt_Com_RazSoc.Enabled = True
            txt_Com_NomCom.Enabled = True
            cmb_Com_GirCom.Enabled = True
            txt_Com_GirCom.Enabled = True
            cmb_Com_TipVia.Enabled = True
            txt_Com_NomVia.Enabled = True
            txt_Com_Numero.Enabled = True
            txt_Com_Interi.Enabled = True
            cmb_Com_TipZon.Enabled = True
            txt_Com_NomZon.Enabled = True
            cmb_Com_DptDir.Enabled = True
            cmb_Com_PrvDir.Enabled = True
            cmb_Com_DstDir.Enabled = True
            txt_Com_Refere.Enabled = True
            txt_Com_Telef1.Enabled = True
            txt_Com_Telef2.Enabled = True
            txt_Com_NumFax.Enabled = True
            
            txt_Com_RazSoc.Text = ""
            txt_Com_NomCom.Text = ""

            cmb_Com_GirCom.ListIndex = -1
            txt_Com_GirCom.Text = ""
            cmb_Com_TipVia.ListIndex = -1
            txt_Com_NomVia.Text = ""
            txt_Com_Numero.Text = ""
            txt_Com_Interi.Text = ""
            cmb_Com_TipZon.ListIndex = -1
            txt_Com_NomZon.Text = ""
            cmb_Com_DptDir.ListIndex = -1
            cmb_Com_PrvDir.Clear
            cmb_Com_DstDir.Clear
            txt_Com_Refere.Text = ""
            txt_Com_Telef1.Text = ""
            txt_Com_Telef2.Text = ""
            txt_Com_NumFax.Text = ""
                        
            Call gs_SetFocus(txt_Com_RazSoc)
            
         Case 41
            pnl_Acc_FlgEmp.Caption = "NR"
      
            txt_Acc_RazSoc.Enabled = True
            txt_Acc_NomCom.Enabled = True
            cmb_Acc_GirCom.Enabled = True
            txt_Acc_GirCom.Enabled = True
            cmb_Acc_TipVia.Enabled = True
            txt_Acc_NomVia.Enabled = True
            txt_Acc_Numero.Enabled = True
            txt_Acc_Interi.Enabled = True
            cmb_Acc_TipZon.Enabled = True
            txt_Acc_NomZon.Enabled = True
            cmb_Acc_DptDir.Enabled = True
            cmb_Acc_PrvDir.Enabled = True
            cmb_Acc_DstDir.Enabled = True
            txt_Acc_Refere.Enabled = True
            txt_Acc_Telef1.Enabled = True
            txt_Acc_Telef2.Enabled = True
            txt_Acc_NumFax.Enabled = True
            
            txt_Acc_RazSoc.Text = ""
            txt_Acc_NomCom.Text = ""
            cmb_Acc_GirCom.ListIndex = -1
            txt_Acc_GirCom.Text = ""
            cmb_Acc_TipVia.ListIndex = -1
            txt_Acc_NomVia.Text = ""
            txt_Acc_Numero.Text = ""
            txt_Acc_Interi.Text = ""
            cmb_Acc_TipZon.ListIndex = -1
            txt_Acc_NomZon.Text = ""
            cmb_Acc_DptDir.ListIndex = -1
            cmb_Acc_PrvDir.Clear
            cmb_Acc_DstDir.Clear
            txt_Acc_Refere.Text = ""
            txt_Acc_Telef1.Text = ""
            txt_Acc_Telef2.Text = ""
            txt_Acc_NumFax.Text = ""
            
            Call gs_SetFocus(txt_Acc_RazSoc)
      
      End Select
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   
   Select Case cmb_ActEco.ItemData(cmb_ActEco.ListIndex)
      Case 11
         txt_Dep_RazSoc.Text = Trim(g_rst_Princi!DATGEN_RAZSOC)
         txt_Dep_NomCom.Text = Trim(g_rst_Princi!DATGEN_NOMCOM)
         
         cmb_Dep_GirCom.ListIndex = gf_Busca_Arregl(l_arr_Dep_GirCom, Trim(g_rst_Princi!DATGEN_GCOMCO)) - 1
         txt_Dep_GirCom.Text = Trim(g_rst_Princi!DATGEN_GCOMNO & "")
         
         Call gs_BuscarCombo_Item(cmb_Dep_TipVia, CInt(Trim(g_rst_Princi!DatGen_TipVia)))
         txt_Dep_NomVia.Text = Trim(g_rst_Princi!DatGen_NomVia & "")
         txt_Dep_Numero.Text = Trim(g_rst_Princi!DatGen_Numero & "")
         txt_Dep_Interi.Text = Trim(g_rst_Princi!DatGen_IntDpt & "")
         Call gs_BuscarCombo_Item(cmb_Dep_TipZon, CInt(Trim(g_rst_Princi!DatGen_TipZon)))
         txt_Dep_NomZon.Text = Trim(g_rst_Princi!DatGen_NomZon & "")
         txt_Dep_Refere.Text = Trim(g_rst_Princi!DatGen_Refere & "")
         
         Call gs_BuscarCombo_Item(cmb_Dep_DptDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 1, 2)))
         
         Call moddat_gs_Carga_Provin(cmb_Dep_PrvDir, Format(cmb_Dep_DptDir.ItemData(cmb_Dep_DptDir.ListIndex), "00"))
         Call gs_BuscarCombo_Item(cmb_Dep_PrvDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 3, 2)))
      
         Call moddat_gs_Carga_Distri(cmb_Dep_DstDir, Format(cmb_Dep_DptDir.ItemData(cmb_Dep_DptDir.ListIndex), "00"), Format(cmb_Dep_PrvDir.ItemData(cmb_Dep_PrvDir.ListIndex), "00"))
         Call gs_BuscarCombo_Item(cmb_Dep_DstDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 5, 2)))
         
         txt_Dep_Telef1.Text = Trim(g_rst_Princi!DATGEN_TELEF1 & "")
         txt_Dep_Telef2.Text = Trim(g_rst_Princi!DATGEN_TELEF2 & "")
         txt_Dep_NumFax.Text = Trim(g_rst_Princi!DatGen_NUMFAX & "")
         txt_Dep_TeleRH.Text = Trim(g_rst_Princi!DATGEN_TELERH & "")
         txt_Dep_AnexRH.Text = Trim(g_rst_Princi!DATGEN_ANEXRH & "")
         
         pnl_Dep_FlgEmp.Caption = Format(g_rst_Princi!DATGEN_CLASIF, "00")
         
         txt_Dep_RazSoc.Enabled = False
         txt_Dep_NomCom.Enabled = False
         cmb_Dep_GirCom.Enabled = False
         txt_Dep_GirCom.Enabled = False
         cmb_Dep_TipVia.Enabled = False
         txt_Dep_NomVia.Enabled = False
         txt_Dep_Numero.Enabled = False
         txt_Dep_Interi.Enabled = False
         cmb_Dep_TipZon.Enabled = False
         txt_Dep_NomZon.Enabled = False
         cmb_Dep_DptDir.Enabled = False
         cmb_Dep_PrvDir.Enabled = False
         cmb_Dep_DstDir.Enabled = False
         txt_Dep_Refere.Enabled = False
         txt_Dep_Telef1.Enabled = False
         txt_Dep_Telef2.Enabled = False
         txt_Dep_NumFax.Enabled = False
         txt_Dep_TeleRH.Enabled = False
         txt_Dep_AnexRH.Enabled = False
         
         Call gs_SetFocus(chk_Dep_Sucurs)
         
      Case 21
         txt_Ind_RazSoc.Text = Trim(g_rst_Princi!DATGEN_RAZSOC & "")
         txt_Ind_Tl1Emp.Text = Trim(g_rst_Princi!DATGEN_TELEF1 & "")
         txt_Ind_Tl2Emp.Text = Trim(g_rst_Princi!DATGEN_TELEF2 & "")
         
         pnl_Ind_FlgEmp.Caption = Format(g_rst_Princi!DATGEN_CLASIF, "00")
      
         txt_Ind_RazSoc.Enabled = False
         txt_Ind_Tl1Emp.Enabled = False
         txt_Ind_Tl2Emp.Enabled = False
         
         Call gs_SetFocus(cmb_Ind_NomCar)
      
      Case 31
         txt_Com_RazSoc.Text = Trim(g_rst_Princi!DATGEN_RAZSOC & "")
         txt_Com_NomCom.Text = Trim(g_rst_Princi!DATGEN_NOMCOM & "")
         
         cmb_Com_GirCom.ListIndex = gf_Busca_Arregl(l_arr_Acc_GirCom, Trim(g_rst_Princi!DATGEN_GCOMCO & "")) - 1
         txt_Com_GirCom.Text = Trim(g_rst_Princi!DATGEN_GCOMNO & "")
         
         Call gs_BuscarCombo_Item(cmb_Com_TipVia, CInt(Trim(g_rst_Princi!DatGen_TipVia)))
         txt_Com_NomVia.Text = Trim(g_rst_Princi!DatGen_NomVia & "")
         txt_Com_Numero.Text = Trim(g_rst_Princi!DatGen_Numero & "")
         txt_Com_Interi.Text = Trim(g_rst_Princi!DatGen_IntDpt & "")
         Call gs_BuscarCombo_Item(cmb_Com_TipZon, CInt(Trim(g_rst_Princi!DatGen_TipZon)))
         txt_Com_NomZon.Text = Trim(g_rst_Princi!DatGen_NomZon & "")
         txt_Com_Refere.Text = Trim(g_rst_Princi!DatGen_Refere & "")
         
         Call gs_BuscarCombo_Item(cmb_Com_DptDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 1, 2)))
         
         Call moddat_gs_Carga_Provin(cmb_Com_PrvDir, Format(cmb_Com_DptDir.ItemData(cmb_Com_DptDir.ListIndex), "00"))
         Call gs_BuscarCombo_Item(cmb_Com_PrvDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 3, 2)))
      
         Call moddat_gs_Carga_Distri(cmb_Com_DstDir, Format(cmb_Com_DptDir.ItemData(cmb_Com_DptDir.ListIndex), "00"), Format(cmb_Com_PrvDir.ItemData(cmb_Com_PrvDir.ListIndex), "00"))
         Call gs_BuscarCombo_Item(cmb_Com_DstDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 5, 2)))
         
         txt_Com_Telef1.Text = Trim(g_rst_Princi!DATGEN_TELEF1 & "")
         txt_Com_Telef2.Text = Trim(g_rst_Princi!DATGEN_TELEF2 & "")
         txt_Com_NumFax.Text = Trim(g_rst_Princi!DatGen_NUMFAX & "")
         
         pnl_Com_FlgEmp.Caption = Format(g_rst_Princi!DATGEN_CLASIF, "00")
      
         txt_Com_RazSoc.Enabled = False
         txt_Com_NomCom.Enabled = False
         cmb_Com_GirCom.Enabled = False
         txt_Com_GirCom.Enabled = False
         cmb_Com_TipVia.Enabled = False
         txt_Com_NomVia.Enabled = False
         txt_Com_Numero.Enabled = False
         txt_Com_Interi.Enabled = False
         cmb_Com_TipZon.Enabled = False
         txt_Com_NomZon.Enabled = False
         cmb_Com_DptDir.Enabled = False
         cmb_Com_PrvDir.Enabled = False
         cmb_Com_DstDir.Enabled = False
         txt_Com_Refere.Enabled = False
         txt_Com_Telef1.Enabled = False
         txt_Com_Telef2.Enabled = False
         txt_Com_NumFax.Enabled = False
      
         Call gs_SetFocus(ipp_Com_IngNet)
         
      Case 41
         txt_Acc_RazSoc.Text = Trim(g_rst_Princi!DATGEN_RAZSOC & "")
         txt_Acc_NomCom.Text = Trim(g_rst_Princi!DATGEN_NOMCOM & "")
         
         cmb_Acc_GirCom.ListIndex = gf_Busca_Arregl(l_arr_Acc_GirCom, Trim(g_rst_Princi!DATGEN_GCOMCO)) - 1
         txt_Acc_GirCom.Text = Trim(g_rst_Princi!DATGEN_GCOMNO & "")
         
         Call gs_BuscarCombo_Item(cmb_Acc_TipVia, CInt(Trim(g_rst_Princi!DatGen_TipVia)))
         txt_Acc_NomVia.Text = Trim(g_rst_Princi!DatGen_NomVia & "")
         txt_Acc_Numero.Text = Trim(g_rst_Princi!DatGen_Numero & "")
         txt_Acc_Interi.Text = Trim(g_rst_Princi!DatGen_IntDpt & "")
         Call gs_BuscarCombo_Item(cmb_Acc_TipZon, CInt(Trim(g_rst_Princi!DatGen_TipZon)))
         txt_Acc_NomZon.Text = Trim(g_rst_Princi!DatGen_NomZon & "")
         txt_Acc_Refere.Text = Trim(g_rst_Princi!DatGen_Refere & "")
         
         Call gs_BuscarCombo_Item(cmb_Acc_DptDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 1, 2)))
         
         Call moddat_gs_Carga_Provin(cmb_Acc_PrvDir, Format(cmb_Acc_DptDir.ItemData(cmb_Acc_DptDir.ListIndex), "00"))
         Call gs_BuscarCombo_Item(cmb_Acc_PrvDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 3, 2)))
      
         Call moddat_gs_Carga_Distri(cmb_Acc_DstDir, Format(cmb_Acc_DptDir.ItemData(cmb_Acc_DptDir.ListIndex), "00"), Format(cmb_Acc_PrvDir.ItemData(cmb_Acc_PrvDir.ListIndex), "00"))
         Call gs_BuscarCombo_Item(cmb_Acc_DstDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 5, 2)))
         
         txt_Acc_Telef1.Text = Trim(g_rst_Princi!DATGEN_TELEF1 & "")
         txt_Acc_Telef2.Text = Trim(g_rst_Princi!DATGEN_TELEF2 & "")
         txt_Acc_NumFax.Text = Trim(g_rst_Princi!DatGen_NUMFAX & "")
         
         pnl_Acc_FlgEmp.Caption = Format(g_rst_Princi!DATGEN_CLASIF, "00")
   
         txt_Acc_RazSoc.Enabled = False
         txt_Acc_NomCom.Enabled = False
         cmb_Acc_GirCom.Enabled = False
         txt_Acc_GirCom.Enabled = False
         cmb_Acc_TipVia.Enabled = False
         txt_Acc_NomVia.Enabled = False
         txt_Acc_Numero.Enabled = False
         txt_Acc_Interi.Enabled = False
         cmb_Acc_TipZon.Enabled = False
         txt_Acc_NomZon.Enabled = False
         cmb_Acc_DptDir.Enabled = False
         cmb_Acc_PrvDir.Enabled = False
         cmb_Acc_DstDir.Enabled = False
         txt_Acc_Refere.Enabled = False
         txt_Acc_Telef1.Enabled = False
         txt_Acc_Telef2.Enabled = False
         txt_Acc_NumFax.Enabled = False
         
         Call gs_SetFocus(ipp_Acc_IngNet)
   End Select

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
