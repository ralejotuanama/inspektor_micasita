VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_IngSol_09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form8"
   ClientHeight    =   10005
   ClientLeft      =   2130
   ClientTop       =   1830
   ClientWidth     =   11655
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10005
      Left            =   0
      TabIndex        =   0
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   825
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   645
            Left            =   630
            TabIndex        =   2
            Top             =   60
            Width           =   3195
            _Version        =   65536
            _ExtentX        =   5636
            _ExtentY        =   1138
            _StockProps     =   15
            Caption         =   "Actividades Económicas del Cónyuge"
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
            TabIndex        =   3
            Top             =   30
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
         Begin Threed.SSPanel pnl_Conyug 
            Height          =   405
            Left            =   3720
            TabIndex        =   4
            Top             =   330
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
            Picture         =   "AteCli_frm_010.frx":0000
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   9045
         Left            =   30
         TabIndex        =   5
         Top             =   900
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   15954
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
            Left            =   8130
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   1740
            Width           =   3315
         End
         Begin VB.TextBox txt_Telefo 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   5040
            Width           =   1640
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   32
            Text            =   "cmb_DstDir"
            Top             =   4710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   31
            Text            =   "cmb_DptDir"
            Top             =   4380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   4050
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   2010
            MaxLength       =   120
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   3720
            Width           =   3315
         End
         Begin VB.ComboBox cmb_GirCom 
            Height          =   315
            Left            =   2010
            TabIndex        =   28
            Text            =   "cmb_GirCom"
            Top             =   3060
            Width           =   3315
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   2010
            MaxLength       =   11
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   2400
            Width           =   3315
         End
         Begin VB.ComboBox cmb_ActEco 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1740
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   2070
            Width           =   3315
         End
         Begin VB.TextBox txt_GirCom 
            Height          =   315
            Left            =   8130
            MaxLength       =   250
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   3060
            Width           =   3315
         End
         Begin VB.TextBox txt_RazSoc 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   2730
            Width           =   3315
         End
         Begin VB.TextBox txt_NumFax 
            Height          =   315
            Left            =   8130
            MaxLength       =   12
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   5040
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8130
            MaxLength       =   250
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   4710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8130
            TabIndex        =   20
            Text            =   "cmb_PrvDir"
            Top             =   4380
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8130
            MaxLength       =   120
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   4050
            Width           =   3315
         End
         Begin VB.TextBox txt_Interi 
            Height          =   315
            Left            =   9810
            MaxLength       =   15
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   3720
            Width           =   1640
         End
         Begin VB.TextBox txt_Numero 
            Height          =   315
            Left            =   8130
            MaxLength       =   15
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   3720
            Width           =   1640
         End
         Begin VB.ComboBox cmb_CodCiu 
            Height          =   315
            Left            =   8130
            Sorted          =   -1  'True
            TabIndex        =   16
            Text            =   "cmb_CodCiu"
            Top             =   2400
            Width           =   3315
         End
         Begin VB.CommandButton cmd_NueAct 
            Caption         =   "&Nueva Actividad"
            Height          =   375
            Left            =   9660
            TabIndex        =   15
            Top             =   330
            Width           =   1755
         End
         Begin VB.CommandButton cmd_BorAct 
            Caption         =   "&Borrar Actividad"
            Height          =   375
            Left            =   9660
            TabIndex        =   14
            Top             =   750
            Width           =   1755
         End
         Begin VB.CommandButton cmd_EdiAct 
            Caption         =   "&Editar Actividad"
            Height          =   375
            Left            =   9660
            TabIndex        =   13
            Top             =   1170
            Width           =   1755
         End
         Begin VB.CommandButton cmd_Agrega 
            Caption         =   "&Agregar a Lista"
            Height          =   375
            Left            =   60
            TabIndex        =   12
            Top             =   7680
            Width           =   1755
         End
         Begin VB.CommandButton cmd_Cancel 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   1830
            TabIndex        =   11
            Top             =   7680
            Width           =   1755
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10050
            Picture         =   "AteCli_frm_010.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Grabar Datos"
            Top             =   8310
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10770
            Picture         =   "AteCli_frm_010.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   8310
            Width           =   675
         End
         Begin VB.TextBox txt_Sucurs 
            Height          =   315
            Left            =   8130
            MaxLength       =   250
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   2730
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   3390
            Width           =   3315
         End
         Begin VB.TextBox txt_Telef1 
            Height          =   315
            Left            =   3660
            MaxLength       =   12
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   5040
            Width           =   1640
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1245
            Left            =   30
            TabIndex        =   34
            Top             =   330
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   2196
            _Version        =   393216
            Rows            =   12
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_FlgEmp 
            Height          =   315
            Left            =   5370
            TabIndex        =   35
            ToolTipText     =   "Empresa No Registrada"
            Top             =   2400
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
            BevelInner      =   2
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   3300
            TabIndex        =   36
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
            TabIndex        =   37
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   90
            Left            =   30
            TabIndex        =   38
            Top             =   1620
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   90
            Left            =   30
            TabIndex        =   39
            Top             =   5400
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   90
            Left            =   30
            TabIndex        =   40
            Top             =   8100
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   90
            Left            =   30
            TabIndex        =   41
            Top             =   7530
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
         Begin Threed.SSPanel pnl_Rentis 
            Height          =   2025
            Left            =   60
            TabIndex        =   43
            Top             =   5490
            Width           =   11385
            _Version        =   65536
            _ExtentX        =   20082
            _ExtentY        =   3572
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
            BevelOuter      =   0
            Begin EditLib.fpDoubleSingle ipp_Ren_IngNet 
               Height          =   315
               Left            =   1950
               TabIndex        =   44
               Top             =   30
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin VB.Label Label10 
               Caption         =   "Ingreso Neto S/."
               Height          =   285
               Left            =   60
               TabIndex        =   45
               Top             =   30
               Width           =   1965
            End
         End
         Begin Threed.SSPanel pnl_Accion 
            Height          =   2025
            Left            =   60
            TabIndex        =   46
            Top             =   5490
            Width           =   11385
            _Version        =   65536
            _ExtentX        =   20082
            _ExtentY        =   3572
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
            BevelOuter      =   0
            Begin EditLib.fpDoubleSingle ipp_Acc_IngNet 
               Height          =   315
               Left            =   1950
               TabIndex        =   47
               Top             =   30
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpLongInteger ipp_Acc_UtiVec 
               Height          =   315
               Left            =   1950
               TabIndex        =   48
               Top             =   360
               Width           =   525
               _Version        =   196608
               _ExtentX        =   926
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpLongInteger ipp_Acc_UtiAno 
               Height          =   315
               Left            =   3540
               TabIndex        =   49
               Top             =   360
               Width           =   525
               _Version        =   196608
               _ExtentX        =   926
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_Acc_Porcen 
               Height          =   315
               Left            =   1950
               TabIndex        =   50
               Top             =   690
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin VB.Label Label53 
               Caption         =   "veces cada"
               Height          =   285
               Left            =   2520
               TabIndex        =   55
               Top             =   390
               Width           =   1035
            End
            Begin VB.Label Label51 
               Caption         =   "Año(s)"
               Height          =   285
               Left            =   4110
               TabIndex        =   54
               Top             =   390
               Width           =   675
            End
            Begin VB.Label Label50 
               Caption         =   "Particip. Utilidades:"
               Height          =   285
               Left            =   60
               TabIndex        =   53
               Top             =   360
               Width           =   1965
            End
            Begin VB.Label Label52 
               Caption         =   "Ingreso Neto S/."
               Height          =   285
               Left            =   60
               TabIndex        =   52
               Top             =   30
               Width           =   1965
            End
            Begin VB.Label Label49 
               Caption         =   "Porcentaje Accionariado:"
               Height          =   315
               Left            =   60
               TabIndex        =   51
               Top             =   690
               Width           =   1905
            End
         End
         Begin Threed.SSPanel pnl_TraDep 
            Height          =   2025
            Left            =   60
            TabIndex        =   56
            Top             =   5490
            Width           =   11355
            _Version        =   65536
            _ExtentX        =   20029
            _ExtentY        =   3572
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
            BevelOuter      =   0
            Begin VB.ComboBox cmb_Dep_FreHab 
               Height          =   315
               Left            =   8070
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   30
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Dep_NomCar 
               Height          =   315
               Left            =   1950
               TabIndex        =   67
               Text            =   "cmb_Dep_NomCar"
               Top             =   360
               Width           =   3315
            End
            Begin VB.TextBox txt_Dep_NomCar 
               Height          =   315
               Left            =   8070
               MaxLength       =   250
               TabIndex        =   66
               Text            =   "Text1"
               Top             =   360
               Width           =   3315
            End
            Begin VB.TextBox txt_Dep_Telefo 
               Height          =   315
               Left            =   1950
               MaxLength       =   12
               TabIndex        =   65
               Text            =   "Text1"
               Top             =   1020
               Width           =   1640
            End
            Begin VB.TextBox txt_Dep_NumAnx 
               Height          =   315
               Left            =   3600
               MaxLength       =   5
               TabIndex        =   64
               Text            =   "Text1"
               Top             =   1020
               Width           =   1640
            End
            Begin VB.TextBox txt_Dep_TlfRhh 
               Height          =   315
               Left            =   8070
               MaxLength       =   12
               TabIndex        =   63
               Text            =   "Text1"
               Top             =   1350
               Width           =   1640
            End
            Begin VB.TextBox txt_Dep_AnxRhh 
               Height          =   315
               Left            =   9720
               MaxLength       =   5
               TabIndex        =   62
               Text            =   "Text1"
               Top             =   1350
               Width           =   1640
            End
            Begin VB.CommandButton cmd_ActEco 
               Height          =   675
               Left            =   5100
               Picture         =   "AteCli_frm_010.frx":0A56
               Style           =   1  'Graphical
               TabIndex        =   61
               ToolTipText     =   "Actividades Económicas"
               Top             =   5670
               Width           =   675
            End
            Begin VB.TextBox txt_Dep_DirEle 
               Height          =   315
               Left            =   1950
               MaxLength       =   120
               TabIndex        =   60
               Text            =   "Text1"
               Top             =   1350
               Width           =   1640
            End
            Begin VB.CheckBox chk_Dep_DirEle 
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
               Left            =   3630
               TabIndex        =   59
               Top             =   1380
               Width           =   1485
            End
            Begin VB.TextBox txt_Dep_NomAre 
               Height          =   315
               Left            =   1950
               MaxLength       =   250
               TabIndex        =   58
               Text            =   "Text1"
               Top             =   690
               Width           =   3315
            End
            Begin VB.TextBox txt_Dep_Celula 
               Height          =   315
               Left            =   8070
               MaxLength       =   12
               TabIndex        =   57
               Text            =   "Text1"
               Top             =   1020
               Width           =   1640
            End
            Begin EditLib.fpDoubleSingle ipp_Dep_IngNet 
               Height          =   315
               Left            =   1950
               TabIndex        =   69
               Top             =   30
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDateTime ipp_Dep_FecIng 
               Height          =   315
               Left            =   8070
               TabIndex        =   70
               Top             =   690
               Width           =   1640
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
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
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               StartMonth      =   4
               ButtonAlign     =   0
               BoundDataType   =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDateTime ipp_Dep_FecCes 
               Height          =   315
               Left            =   1950
               TabIndex        =   71
               Top             =   1680
               Width           =   1640
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
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
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               StartMonth      =   4
               ButtonAlign     =   0
               BoundDataType   =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin VB.Label Label35 
               Caption         =   "Ingreso Neto S/."
               Height          =   285
               Left            =   60
               TabIndex        =   82
               Top             =   30
               Width           =   1485
            End
            Begin VB.Label Label36 
               Caption         =   "Frecuencia Haberes:"
               Height          =   315
               Left            =   6060
               TabIndex        =   81
               Top             =   30
               Width           =   1905
            End
            Begin VB.Label Label37 
               Caption         =   "Cargo:"
               Height          =   285
               Left            =   60
               TabIndex        =   80
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label Label38 
               Caption         =   "Fecha de Ingreso:"
               Height          =   315
               Left            =   6060
               TabIndex        =   79
               Top             =   690
               Width           =   1905
            End
            Begin VB.Label Label39 
               Caption         =   "Teléfono / Anexo:"
               Height          =   285
               Left            =   60
               TabIndex        =   78
               Top             =   1020
               Width           =   1575
            End
            Begin VB.Label Label40 
               Caption         =   "Telf. / Anx (RR.HH):"
               Height          =   285
               Left            =   6060
               TabIndex        =   77
               Top             =   1350
               Width           =   2055
            End
            Begin VB.Label Label3 
               Caption         =   "Fecha de Cese:"
               Height          =   315
               Left            =   60
               TabIndex        =   76
               Top             =   1680
               Width           =   1905
            End
            Begin VB.Label Label6 
               Caption         =   "Cargo (Especificar):"
               Height          =   285
               Left            =   6060
               TabIndex        =   75
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label Label17 
               Caption         =   "E-mail:"
               Height          =   285
               Left            =   60
               TabIndex        =   74
               Top             =   1350
               Width           =   1485
            End
            Begin VB.Label Label2 
               Caption         =   "Area:"
               Height          =   285
               Left            =   60
               TabIndex        =   73
               Top             =   690
               Width           =   1605
            End
            Begin VB.Label Label8 
               Caption         =   "Celular:"
               Height          =   285
               Left            =   6060
               TabIndex        =   72
               Top             =   1020
               Width           =   1575
            End
         End
         Begin Threed.SSPanel pnl_TraInd 
            Height          =   2025
            Left            =   60
            TabIndex        =   83
            Top             =   5490
            Width           =   11415
            _Version        =   65536
            _ExtentX        =   20135
            _ExtentY        =   3572
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
            BevelOuter      =   0
            Begin VB.TextBox txt_Ind_NumDoc 
               Height          =   315
               Left            =   1950
               MaxLength       =   11
               TabIndex        =   89
               Text            =   "Text1"
               Top             =   1020
               Width           =   2715
            End
            Begin VB.ComboBox cmb_Ind_ConLoc 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   88
               Top             =   360
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Ind_TipDoc 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   87
               Top             =   690
               Width           =   3315
            End
            Begin VB.TextBox txt_Ind_RazSoc 
               Height          =   315
               Left            =   1950
               MaxLength       =   250
               TabIndex        =   86
               Text            =   "Text1"
               Top             =   1350
               Width           =   9285
            End
            Begin VB.TextBox txt_Ind_Telef1 
               Height          =   315
               Left            =   1950
               MaxLength       =   12
               TabIndex        =   85
               Text            =   "Text1"
               Top             =   1680
               Width           =   1275
            End
            Begin VB.TextBox txt_Ind_Telef2 
               Height          =   315
               Left            =   3270
               MaxLength       =   12
               TabIndex        =   84
               Text            =   "Text1"
               Top             =   1680
               Width           =   1275
            End
            Begin EditLib.fpDoubleSingle ipp_Ind_IngNet 
               Height          =   315
               Left            =   1950
               TabIndex        =   90
               Top             =   30
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin VB.Label Label54 
               Caption         =   "Razón Social:"
               Height          =   315
               Left            =   60
               TabIndex        =   96
               Top             =   1350
               Width           =   1905
            End
            Begin VB.Label Label56 
               Caption         =   "Nro. Docum. Empresa:"
               Height          =   285
               Left            =   60
               TabIndex        =   95
               Top             =   1020
               Width           =   2055
            End
            Begin VB.Label Label58 
               Caption         =   "Tipo Docum. Empresa:"
               Height          =   315
               Left            =   60
               TabIndex        =   94
               Top             =   690
               Width           =   1905
            End
            Begin VB.Label Label59 
               Caption         =   "Contrato Locación:"
               Height          =   285
               Left            =   60
               TabIndex        =   93
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label Label61 
               Caption         =   "Ingreso Neto Mensual S/.:"
               Height          =   285
               Left            =   60
               TabIndex        =   92
               Top             =   30
               Width           =   2025
            End
            Begin VB.Label Label4 
               Caption         =   "Teléfonos:"
               Height          =   285
               Left            =   60
               TabIndex        =   91
               Top             =   1680
               Width           =   1485
            End
         End
         Begin Threed.SSPanel pnl_Comerc 
            Height          =   2025
            Left            =   60
            TabIndex        =   97
            Top             =   5490
            Width           =   11385
            _Version        =   65536
            _ExtentX        =   20082
            _ExtentY        =   3572
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
            BevelOuter      =   0
            Begin VB.ComboBox cmb_Com_RegTri 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   102
               Top             =   690
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Com_TipLoc 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   101
               Top             =   1050
               Width           =   3315
            End
            Begin VB.TextBox txt_Com_Tl1Arr 
               Height          =   315
               Left            =   8070
               MaxLength       =   12
               TabIndex        =   100
               Text            =   "Text1"
               Top             =   1380
               Width           =   1640
            End
            Begin VB.TextBox txt_Com_Tl2Arr 
               Height          =   315
               Left            =   9720
               MaxLength       =   5
               TabIndex        =   99
               Text            =   "Text1"
               Top             =   1380
               Width           =   1640
            End
            Begin VB.TextBox txt_Com_NomArr 
               Height          =   315
               Left            =   1950
               MaxLength       =   250
               TabIndex        =   98
               Text            =   "Text1"
               Top             =   1380
               Width           =   3315
            End
            Begin EditLib.fpDoubleSingle ipp_Com_IngNet 
               Height          =   315
               Left            =   1950
               TabIndex        =   103
               Top             =   30
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDateTime ipp_Com_FecIni 
               Height          =   315
               Left            =   8070
               TabIndex        =   104
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
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
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               StartMonth      =   4
               ButtonAlign     =   0
               BoundDataType   =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_Com_VtaMen 
               Height          =   315
               Left            =   1950
               TabIndex        =   105
               Top             =   360
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_Com_AlqMen 
               Height          =   315
               Left            =   8070
               TabIndex        =   106
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_Com_PorPar 
               Height          =   315
               Left            =   8070
               TabIndex        =   107
               Top             =   720
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
               ThreeDInsideHighlightColor=   -2147483633
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
               ThreeDTextHighlightColor=   -2147483633
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin VB.Label Label45 
               Caption         =   "Régimen Tributario:"
               Height          =   315
               Left            =   60
               TabIndex        =   116
               Top             =   690
               Width           =   1905
            End
            Begin VB.Label Label46 
               Caption         =   "Ventas Mensuales S/."
               Height          =   285
               Left            =   60
               TabIndex        =   115
               Top             =   360
               Width           =   1965
            End
            Begin VB.Label Label47 
               Caption         =   "Ingreso Neto Mensual S/."
               Height          =   285
               Left            =   60
               TabIndex        =   114
               Top             =   30
               Width           =   1965
            End
            Begin VB.Label Label48 
               Caption         =   "Fecha de Inicio Operac.:"
               Height          =   315
               Left            =   6060
               TabIndex        =   113
               Top             =   390
               Width           =   1905
            End
            Begin VB.Label Label11 
               Caption         =   "Tipo Local Comercial:"
               Height          =   315
               Left            =   60
               TabIndex        =   112
               Top             =   1050
               Width           =   1905
            End
            Begin VB.Label Label42 
               Caption         =   "Teléfonos:"
               Height          =   285
               Left            =   6060
               TabIndex        =   111
               Top             =   1380
               Width           =   2055
            End
            Begin VB.Label Label12 
               Caption         =   "Nombre Arrendador:"
               Height          =   285
               Left            =   60
               TabIndex        =   110
               Top             =   1380
               Width           =   1485
            End
            Begin VB.Label Label13 
               Caption         =   "Monto Alq. Mens. US$:"
               Height          =   285
               Left            =   6060
               TabIndex        =   109
               Top             =   1050
               Width           =   1785
            End
            Begin VB.Label Label1 
               Caption         =   "% Participación:"
               Height          =   285
               Left            =   6060
               TabIndex        =   108
               Top             =   720
               Width           =   1785
            End
         End
         Begin VB.Label Label33 
            Caption         =   "Giro Comercial:"
            Height          =   285
            Left            =   90
            TabIndex        =   136
            Top             =   3060
            Width           =   2055
         End
         Begin VB.Label Label32 
            Caption         =   "Razón Social:"
            Height          =   285
            Left            =   90
            TabIndex        =   135
            Top             =   2730
            Width           =   1485
         End
         Begin VB.Label Label31 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Left            =   90
            TabIndex        =   134
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label30 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Left            =   90
            TabIndex        =   133
            Top             =   2070
            Width           =   2055
         End
         Begin VB.Label Label27 
            Caption         =   "Teléfono:"
            Height          =   285
            Left            =   90
            TabIndex        =   132
            Top             =   5040
            Width           =   1485
         End
         Begin VB.Label Label26 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   90
            TabIndex        =   131
            Top             =   4710
            Width           =   1905
         End
         Begin VB.Label Label24 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   90
            TabIndex        =   130
            Top             =   4380
            Width           =   1905
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   90
            TabIndex        =   129
            Top             =   4050
            Width           =   1905
         End
         Begin VB.Label Label20 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   90
            TabIndex        =   128
            Top             =   3720
            Width           =   1485
         End
         Begin VB.Label Label19 
            Caption         =   "Actividad Económica:"
            Height          =   315
            Left            =   90
            TabIndex        =   127
            Top             =   1740
            Width           =   1905
         End
         Begin VB.Label Label34 
            Caption         =   "CIIU"
            Height          =   285
            Left            =   6120
            TabIndex        =   126
            Top             =   2400
            Width           =   1995
         End
         Begin VB.Label Label29 
            Caption         =   "Fax:"
            Height          =   285
            Left            =   6120
            TabIndex        =   125
            Top             =   5040
            Width           =   1485
         End
         Begin VB.Label Label18 
            Caption         =   "Orden Actividad Econom.:"
            Height          =   315
            Left            =   6120
            TabIndex        =   124
            Top             =   1740
            Width           =   2115
         End
         Begin VB.Label Label28 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   6120
            TabIndex        =   123
            Top             =   4710
            Width           =   1485
         End
         Begin VB.Label Label25 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   6120
            TabIndex        =   122
            Top             =   4380
            Width           =   1905
         End
         Begin VB.Label Label23 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   6120
            TabIndex        =   121
            Top             =   4050
            Width           =   1485
         End
         Begin VB.Label Label21 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   6120
            TabIndex        =   120
            Top             =   3720
            Width           =   2055
         End
         Begin VB.Label Label41 
            Caption         =   "Sucursal:"
            Height          =   285
            Left            =   6120
            TabIndex        =   119
            Top             =   2730
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Giro Comercial (Especif.):"
            Height          =   285
            Left            =   6120
            TabIndex        =   118
            Top             =   3060
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Vía:"
            Height          =   285
            Left            =   90
            TabIndex        =   117
            Top             =   3390
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "frm_IngSol_09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_GirCom()      As moddat_tpo_Genera
Dim l_arr_Dep_NomCar()  As moddat_tpo_Genera
Dim l_int_FlgCmb        As Integer
Dim l_str_DstDir        As String
Dim l_str_DptDir        As String
Dim l_str_PrvDir        As String
Dim l_str_CodCiu        As String
Dim l_str_GirCom        As String
Dim l_str_NomCar        As String
Dim l_int_OrdAct        As Integer
Dim l_int_FlgGrb        As Integer

Private Sub cmb_ActEco_Click()
   If cmb_ActEco.ListIndex <> -1 Then
      Select Case cmb_ActEco.ItemData(cmb_ActEco.ListIndex)
         Case 11, 12
            pnl_TraDep.Visible = True
            pnl_TraInd.Visible = False
            pnl_Comerc.Visible = False
            pnl_Accion.Visible = False
            pnl_Rentis.Visible = False
            
            Call fs_Activa_Dep(True)
            Call fs_Limpia_TraDep
            
            If Len(Trim(txt_NumDoc.Text)) > 0 Then
               pnl_FlgEmp.Visible = True
            Else
               pnl_FlgEmp.Visible = False
            End If
            
            txt_RazSoc.Enabled = True
            
         Case 21
            pnl_TraDep.Visible = False
            pnl_TraInd.Visible = True
            pnl_Comerc.Visible = False
            pnl_Accion.Visible = False
            pnl_Rentis.Visible = False
      
            Call fs_Activa_Ind(True)
            Call fs_Limpia_TraInd
            
            pnl_FlgEmp.Visible = False
            txt_RazSoc.Text = ""
            txt_Sucurs.Text = ""
            txt_RazSoc.Enabled = False
            txt_Sucurs.Enabled = False
         Case 31
            pnl_TraDep.Visible = False
            pnl_TraInd.Visible = False
            pnl_Comerc.Visible = True
            pnl_Accion.Visible = False
            pnl_Rentis.Visible = False
            
            Call fs_Activa_Com(True)
            Call fs_Limpia_TraCom
            
            pnl_FlgEmp.Visible = False
            txt_RazSoc.Enabled = True
            
         Case 41
            pnl_TraDep.Visible = False
            pnl_TraInd.Visible = False
            pnl_Comerc.Visible = False
            pnl_Rentis.Visible = False
            pnl_Accion.Visible = True
            pnl_Rentis.Visible = False
            
            Call fs_Activa_Acc(True)
            Call fs_Limpia_TraAcc
            
            pnl_FlgEmp.Visible = False
            txt_RazSoc.Enabled = True
      
         Case 51
            pnl_TraDep.Visible = False
            pnl_TraInd.Visible = False
            pnl_Comerc.Visible = False
            pnl_Accion.Visible = False
            pnl_Rentis.Visible = True
            
            Call fs_Activa_Ren(True)
            Call fs_Limpia_TraRen
            
            pnl_FlgEmp.Visible = False
            txt_RazSoc.Text = ""
            txt_Sucurs.Text = ""
            txt_RazSoc.Enabled = False
            txt_Sucurs.Enabled = False
      End Select
      
      Call gs_SetFocus(cmb_OrdAct)
   End If
End Sub

Private Sub cmb_ActEco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_OrdAct)
   End If
End Sub

Private Sub cmb_CodCiu_Change()
   l_str_CodCiu = cmb_CodCiu.Text
End Sub

Private Sub cmb_CodCiu_Click()
   If cmb_CodCiu.ListIndex > -1 Then
      If l_int_FlgCmb Then
         If txt_RazSoc.Enabled Then
            Call gs_SetFocus(txt_RazSoc)
         Else
            Call gs_SetFocus(cmb_GirCom)
         End If
      End If
   End If
End Sub

Private Sub cmb_CodCiu_GotFocus()
   l_int_FlgCmb = True
   l_str_CodCiu = cmb_CodCiu.Text
End Sub

Private Sub cmb_CodCiu_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_CodCiu, l_str_CodCiu)
      l_int_FlgCmb = True
      
      If cmb_CodCiu.ListIndex > -1 Then
         l_str_CodCiu = ""
      End If
      
      If txt_RazSoc.Enabled Then
         Call gs_SetFocus(txt_RazSoc)
      Else
         Call gs_SetFocus(cmb_GirCom)
      End If
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

Private Sub cmb_com_TipLoc_Click()
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

Private Sub cmb_com_TipLoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_com_TipLoc_Click
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

Private Sub cmb_Dep_NomCar_Change()
   l_str_NomCar = cmb_Dep_NomCar.Text
End Sub

Private Sub cmb_Dep_NomCar_Click()
   If cmb_Dep_NomCar.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Dep_NomCar)
      End If
   End If
End Sub

Private Sub cmb_Dep_NomCar_GotFocus()
   l_int_FlgCmb = True
   l_str_NomCar = cmb_Dep_NomCar.Text
End Sub

Private Sub cmb_Dep_NomCar_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Dep_NomCar, l_str_NomCar)
      l_int_FlgCmb = True
      
      If cmb_Dep_NomCar.ListIndex > -1 Then
         l_str_NomCar = ""
      End If
      
      Call gs_SetFocus(txt_Dep_NomCar)
   End If
End Sub

Private Sub cmb_GirCom_Change()
   l_str_GirCom = cmb_GirCom.Text
End Sub

Private Sub cmb_GirCom_Click()
   If cmb_GirCom.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_GirCom)
      End If
   End If
End Sub

Private Sub cmb_GirCom_GotFocus()
   l_int_FlgCmb = True
   l_str_GirCom = cmb_GirCom.Text
End Sub

Private Sub cmb_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_GirCom, l_str_GirCom)
      l_int_FlgCmb = True
      
      If cmb_GirCom.ListIndex > -1 Then
         l_str_GirCom = ""
      End If
      
      Call gs_SetFocus(txt_GirCom)
   End If
End Sub

Private Sub cmb_Ind_ConLoc_Click()
   cmb_Ind_TipDoc.Enabled = False
   txt_Ind_NumDoc.Enabled = False
   txt_Ind_RazSoc.Enabled = False
   txt_Ind_Telef1.Enabled = False
   txt_Ind_Telef2.Enabled = False
   
   Call gs_SetFocus(cmd_Agrega)
   
   If cmb_Ind_ConLoc.ListIndex > -1 Then
      If cmb_Ind_ConLoc.ItemData(cmb_Ind_ConLoc.ListIndex) = 1 Then
         cmb_Ind_TipDoc.Enabled = True
         txt_Ind_NumDoc.Enabled = True
         txt_Ind_RazSoc.Enabled = True
         txt_Ind_Telef1.Enabled = True
         txt_Ind_Telef2.Enabled = True
         
         Call gs_SetFocus(cmb_Ind_TipDoc)
      Else
         cmb_Ind_TipDoc.ListIndex = -1
         txt_Ind_NumDoc.Text = ""
         txt_Ind_RazSoc.Text = ""
         txt_Ind_Telef1.Text = ""
         txt_Ind_Telef2.Text = ""
      End If
   End If
End Sub

Private Sub cmb_Ind_ConLoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Ind_ConLoc_Click
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
   If cmb_ActEco.ListIndex > -1 And cmb_OrdAct.ListIndex > -1 Then
      If cmb_OrdAct.ItemData(cmb_OrdAct.ListIndex) = 9 And (cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 11 Or cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 12) Then
         ipp_Dep_FecCes.Enabled = True
      Else
         ipp_Dep_FecCes.Text = Format(Date - CDate(365 * 2), "dd/mm/yyyy")
         ipp_Dep_FecCes.Enabled = False
      End If
      
      Call gs_SetFocus(cmb_TipDoc)
   Else
      If cmb_OrdAct.ListIndex = -1 Then
         Call gs_SetFocus(cmb_TipDoc)
      End If
      
      If cmb_ActEco.ListIndex = -1 Then
         Call gs_SetFocus(cmb_ActEco)
      End If
   End If
End Sub

Private Sub cmb_OrdAct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_OrdAct_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   Dim r_int_Indice     As Integer
   
   'Validando Datos Generales
   If cmb_ActEco.ListIndex = -1 Then
      MsgBox "Seleccione la Actividad Económica.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_ActEco)
      Exit Sub
   End If
   
   If cmb_OrdAct.ListIndex = -1 Then
      MsgBox "Seleccione el Orden de la Actividad Económica.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_OrdAct)
      Exit Sub
   End If
   
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Documento de Identidad.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc.Text)) <> 11 Then
      MsgBox "Ingrese correctamente el Número de Documento de Identidad.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If

   If Not gf_Valida_RUC(Mid(txt_NumDoc.Text, 1, Len(txt_NumDoc.Text) - 1), Right(txt_NumDoc.Text, 1)) Then
      MsgBox "El Número de Documento de Identidad no es válido.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If

   If cmb_CodCiu.ListIndex = -1 Then
      MsgBox "Seleccione el Código CIIU.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_CodCiu)
      Exit Sub
   End If
   
   If cmb_ActEco.ItemData(cmb_ActEco.ListIndex) <> 21 And cmb_ActEco.ItemData(cmb_ActEco.ListIndex) <> 51 Then
      If Len(Trim(txt_RazSoc.Text)) = 0 Then
         MsgBox "Ingrese la Razón Social.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(txt_RazSoc)
         Exit Sub
      End If
   End If
   
   If cmb_GirCom.ListIndex = -1 Then
      MsgBox "Seleccione el Giro Comercial.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_GirCom)
      Exit Sub
   End If

   'Validar Ingreso de Giro Comercial (TXT)
   
   
   If cmb_TipVia.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Vía de la Dirección.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_TipVia)
      Exit Sub
   End If

   If Len(Trim(txt_NomVia.Text)) = 0 Then
      MsgBox "Ingrese el Nombre de la Vóa de la Dirección.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_NomVia)
      Exit Sub
   End If
   
   If Len(Trim(txt_Numero.Text)) = 0 Then
      MsgBox "Ingrese el Número en la Vía de la Dirección.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_Numero)
      Exit Sub
   End If
   
   If cmb_TipZon.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Zona de la Dirección.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_TipZon)
      Exit Sub
   End If

   If cmb_TipZon.ItemData(cmb_TipZon.ListIndex) <> 12 Then
      If Len(Trim(txt_NomZon.Text)) = 0 Then
         MsgBox "Ingrese el Nombre de la Zona de la Dirección.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(txt_NomZon)
         Exit Sub
      End If
   End If

   If cmb_DptDir.ListIndex = -1 Then
      MsgBox "Seleccione el Departamento de la Dirección.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_DptDir)
      Exit Sub
   End If
   
   If cmb_PrvDir.ListIndex = -1 Then
      MsgBox "Seleccione la Provincia de la Dirección.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_PrvDir)
      Exit Sub
   End If
   
   If cmb_DstDir.ListIndex = -1 Then
      MsgBox "Seleccione el Distrito de la Dirección.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_DstDir)
      Exit Sub
   End If
   
   If Len(Trim(txt_Telefo.Text)) = 0 Then
      MsgBox "Ingrese el Teléfono de la Empresa.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_Telefo)
      Exit Sub
   End If
   
   'Validar según Actividad Económica
   If pnl_TraDep.Visible Then
      If Not ff_Valida_TraDep() Then
         Exit Sub
      End If
   End If
   
   If pnl_TraInd.Visible Then
      If Not ff_Valida_TraInd() Then
         Exit Sub
      End If
   End If
   
   If pnl_Comerc.Visible Then
      If Not ff_Valida_Comerc() Then
         Exit Sub
      End If
   End If
   
   If pnl_Accion.Visible Then
      If Not ff_Valida_Accion() Then
         Exit Sub
      End If
   End If
   
   If pnl_Rentis.Visible Then
      If Not ff_Valida_Rentis() Then
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de agregar el Item a la Lista?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_con_AteCli) <> vbYes Then
      Exit Sub
   End If
   
   If l_int_FlgGrb = 1 Then
      r_int_Indice = grd_Listad.Rows + 1
      
      ReDim Preserve modatecli_g_arr_CygActEco(r_int_Indice)
   Else
      r_int_Indice = l_int_OrdAct
   End If
   
   'Limpiando Variables del Arreglo
   Call fs_Limpia_Arreglo(r_int_Indice)
   Call fs_Limpia_Arreglo_TraDep(r_int_Indice)
   Call fs_Limpia_Arreglo_TraInd(r_int_Indice)
   Call fs_Limpia_Arreglo_Comerc(r_int_Indice)
   Call fs_Limpia_Arreglo_Accion(r_int_Indice)
   Call fs_Limpia_Arreglo_Rentis(r_int_Indice)
   
   'Grabando en Arreglo de Memoria
   Call fs_Arreglo_Genera(r_int_Indice)
   
   If pnl_TraDep.Visible Then
      Call fs_Arreglo_TraDep(r_int_Indice)
   End If
   
   If pnl_TraInd.Visible Then
      Call fs_Arreglo_TraInd(r_int_Indice)
   End If
   
   If pnl_Comerc.Visible Then
      Call fs_Arreglo_Comerc(r_int_Indice)
   End If
   
   If pnl_Accion.Visible Then
      Call fs_Arreglo_Accion(r_int_Indice)
   End If
   
   If pnl_Rentis.Visible Then
      Call fs_Arreglo_Rentis(r_int_Indice)
   End If
   
   If l_int_FlgGrb = 1 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
   Else
      grd_Listad.Row = l_int_OrdAct - 1
   End If
   
   grd_Listad.Col = 0:  grd_Listad.Text = cmb_OrdAct.Text
   grd_Listad.Col = 1:  grd_Listad.Text = cmb_ActEco.Text
   grd_Listad.Col = 2:  grd_Listad.Text = cmb_OrdAct.ItemData(cmb_OrdAct.ListIndex)
   
   Call gs_UbiIniGrid(grd_Listad)
   
   Call fs_Activa(False)
   Call fs_Limpia
   
   Call fs_Activa_Dep(False)
   Call fs_Limpia_TraDep
   pnl_TraDep.Visible = True
   
   Call gs_SetFocus(grd_Listad)
   
   cmd_Grabar.Enabled = True
   cmd_EdiAct.Enabled = True
   cmd_BorAct.Enabled = True
End Sub

Private Sub cmd_BorAct_Click()
   Dim r_int_OrdAct  As Integer
   
   grd_Listad.Col = 2
   r_int_OrdAct = CInt(grd_Listad.Text)
         
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro de eliminar la Actividad de la Lista?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
      Exit Sub
   End If

   Call fs_BorIte(grd_Listad.Row + 1)
   
   If grd_Listad.Rows = 1 Then
      grd_Listad.Rows = 0
   Else
      grd_Listad.RemoveItem grd_Listad.Row
   End If
   
   If grd_Listad.Rows = 0 Then
      cmd_Grabar.Enabled = False
      cmd_EdiAct.Enabled = False
      cmd_BorAct.Enabled = False
   Else

      
      Call gs_UbiIniGrid(grd_Listad)
   End If
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Activa(False)
   Call fs_Limpia
   
   Call fs_Activa_Dep(False)
   Call fs_Limpia_TraDep
   pnl_TraDep.Visible = True
   
   If grd_Listad.Rows > 0 Then
      Call gs_SetFocus(grd_Listad)
   
      cmd_Grabar.Enabled = True
      cmd_EdiAct.Enabled = True
      cmd_BorAct.Enabled = True
   Else
      Call gs_SetFocus(cmd_NueAct)
   End If
End Sub

Private Sub cmd_EdiAct_Click()
   l_int_OrdAct = grd_Listad.Row + 1
   
   l_int_FlgCmb = True
   
   Call fs_Activa(True)
   
   'Pasando del Arreglo a los Controles
   Call gs_BuscarCombo_Item(cmb_ActEco, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_ActEco)
   Call gs_BuscarCombo_Item(cmb_OrdAct, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_OrdAct)
   Call gs_BuscarCombo_Item(cmb_TipDoc, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_TipDoc)
   txt_NumDoc.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_NumDoc
   
   If modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_ActEco <> 21 And modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_ActEco <> 51 Then
      Call fs_BusEmp
   End If
   
   Call gs_BuscarCombo_Item(cmb_CodCiu, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_CodCiu)
   txt_RazSoc.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_RazSoc
   txt_Sucurs.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_Sucurs
   cmb_GirCom.ListIndex = gf_Busca_Arregl(l_arr_GirCom, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_GiroCd) - 1
   txt_GirCom.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_GiroNm
   Call gs_BuscarCombo_Item(cmb_TipVia, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_TipVia)
   txt_NomVia.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_NomVia
   txt_Numero.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_Numero
   txt_Interi.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_Interi
   Call gs_BuscarCombo_Item(cmb_TipZon, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_TipZon)
   txt_NomZon.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_NomZon
   Call gs_BuscarCombo_Item(cmb_DptDir, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_DptDir)
   Call gs_BuscarCombo_Item(cmb_PrvDir, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_PrvDir)
   Call gs_BuscarCombo_Item(cmb_DstDir, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_DstDir)
   txt_Refere.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_Refere
   txt_Telefo.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_Telefo
   txt_Telef1.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_Telef1
   txt_NumFax.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_NumFax
   
   'Call fs_Activa(True)
   
   Select Case modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_ActEco
      Case 11, 12
         Call fs_Activa_Dep(True)
         
         ipp_Dep_IngNet.Value = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_IngNet
         Call gs_BuscarCombo_Item(cmb_Dep_FreHab, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_FreHab)
         cmb_Dep_NomCar.ListIndex = gf_Busca_Arregl(l_arr_Dep_NomCar, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_CargoC) - 1
         txt_Dep_NomCar.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_CargoN
         txt_Dep_NomAre.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_NomAre
         ipp_Dep_FecIng.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_FecIng
         txt_Dep_Telefo.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_Telefo
         txt_Dep_NumAnx.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_NAnexo
         txt_Dep_Celula.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_Celula
         txt_Dep_DirEle.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_DirEle
         
         If modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_Autori = 1 Then
            chk_Dep_DirEle.Value = 1
         Else
            chk_Dep_DirEle.Value = 0
         End If
         
         If modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_Autori = 1 Then
            chk_Dep_DirEle.Value = 1
         Else
            chk_Dep_DirEle.Value = 0
         End If
         
         txt_Dep_TlfRhh.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_TlfRhh
         txt_Dep_AnxRhh.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_AnxRhh
         
         If modatecli_g_arr_CygActEco(l_int_OrdAct).ActEco_OrdAct = 9 Then
            ipp_Dep_FecCes.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoDep_FecCes
         End If
         
      Case 21
         Call fs_Activa_Ind(True)
         
         ipp_Ind_IngNet.Value = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoInd_IngNet
         Call gs_BuscarCombo_Item(cmb_Ind_ConLoc, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoInd_ConLoc)
         
         If modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoInd_ConLoc = 1 Then
            Call gs_BuscarCombo_Item(cmb_Ind_TipDoc, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoInd_TipDoc)
            txt_Ind_NumDoc.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoInd_NumDoc
            txt_Ind_RazSoc.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoInd_RazSoc
            txt_Ind_Telef1.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoInd_Telef1
            txt_Ind_Telef2.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoInd_Telef2
         End If
         
      Case 31
         Call fs_Activa_Com(True)
               
         ipp_Com_IngNet.Value = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoCom_IngNet
         ipp_Com_VtaMen.Value = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoCom_VtaMen
         ipp_Com_PorPar.Value = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoCom_PorPar
         ipp_Com_FecIni.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoCom_FecIni
         Call gs_BuscarCombo_Item(cmb_Com_RegTri, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoCom_RegTri)
         Call gs_BuscarCombo_Item(cmb_Com_TipLoc, modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoCom_TipLoc)
         
         If modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoCom_TipLoc = 2 Then
            ipp_Com_AlqMen.Value = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoCom_AlqMen
            txt_Com_NomArr.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoCom_NomArr
            txt_Com_Tl1Arr.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoCom_Tl1Arr
            txt_Com_Tl2Arr.Text = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoCom_Tl2Arr
         End If
         
      Case 41
         Call fs_Activa_Acc(True)
         
         ipp_Acc_IngNet.Value = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoAcc_IngNet
         ipp_Acc_UtiVec.Value = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoAcc_UtiVec
         ipp_Acc_UtiAno.Value = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoAcc_UtiAno
         ipp_Acc_Porcen.Value = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoAcc_Porcen
   
      Case 51
         Call fs_Activa_Ren(True)
         
         ipp_Ren_IngNet.Value = modatecli_g_arr_CygActEco(l_int_OrdAct).ActEcoRen_IngNet
   End Select
   Call gs_SetFocus(cmb_ActEco)
   
   l_int_FlgGrb = 2
   cmd_Grabar.Enabled = False
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_Contad     As Integer
   Dim r_int_OrdAct()   As Integer
   Dim r_int_Linea1     As Integer
   Dim r_int_Linea2     As Integer
   Dim r_int_Linea3     As Integer
   
   r_int_Linea1 = 0
   r_int_Linea2 = 0
   r_int_Linea3 = 0
   
   For r_int_Contad = 1 To UBound(modatecli_g_arr_CygActEco)
      If r_int_Contad = 1 Then
         r_int_Linea1 = modatecli_g_arr_CygActEco(r_int_Contad).ActEco_OrdAct
      End If
      
      If r_int_Contad = 2 Then
         r_int_Linea2 = modatecli_g_arr_CygActEco(r_int_Contad).ActEco_OrdAct
      End If
   
      If r_int_Contad = 3 Then
         r_int_Linea3 = modatecli_g_arr_CygActEco(r_int_Contad).ActEco_OrdAct
      End If
      
      'Obteniendo Actividad Económica Principal - Cliente Cygular
      If modatecli_g_arr_CygActEco(r_int_Contad).ActEco_OrdAct = 1 Then
         modatecli_g_int_ActPri_Cyg = modatecli_g_arr_CygActEco(r_int_Contad).ActEco_ActEco
      End If
   Next r_int_Contad
   
   If r_int_Linea1 <> 1 And r_int_Linea2 <> 1 And r_int_Linea3 <> 1 Then
      MsgBox "Debe registar la Actividad Principal", vbExclamation, modgen_g_con_AteCli
      Exit Sub
   End If
   
   If (r_int_Linea1 = r_int_Linea2 And r_int_Linea1 <> 0 And r_int_Linea2 <> 0) Or (r_int_Linea1 = r_int_Linea3 And r_int_Linea1 <> 0 And r_int_Linea3 <> 0) Or (r_int_Linea2 = r_int_Linea3 And r_int_Linea2 <> 0 And r_int_Linea3 <> 0) Then
      MsgBox "No se puede tener dos Ordenes de Actividad igual.", vbExclamation, modgen_g_con_AteCli
      Exit Sub
   End If
   
   modatecli_g_int_ActEcoCyg = 2
   
   Unload Me
End Sub

Private Sub cmd_NueAct_Click()
   'Cargando Orden de Actividades Económicas
   If grd_Listad.Rows = 3 Then
      MsgBox "No puede ingresar más Actividades Económicas.", vbExclamation, modgen_g_con_AteCli
      Exit Sub
   End If
   
   l_int_FlgGrb = 1
   
   Call fs_Activa(True)
   Call fs_Activa_Dep(False)
   Call fs_Activa_Ind(False)
   Call fs_Activa_Com(False)
   Call fs_Activa_Acc(False)
   Call fs_Activa_Ren(False)
   
   cmd_Grabar.Enabled = False
   
   Call gs_SetFocus(cmb_ActEco)
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("Al salir de esta manera perderá la información ingresada. ¿Está seguro de salir de la ventana?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
      Exit Sub
   End If
   
   If grd_Listad.Rows = 0 Then
      modatecli_g_int_ActEcoCyg = 1
   End If
   
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_int_Contad     As Integer
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_con_AteCli & " Ingreso de Solicitud de Crédito"
   
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Conyug.Caption = CStr(moddat_g_int_CygTDo) & "-" & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom
   
   Call fs_Inicia
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Limpia
   Call fs_Activa(False)
   Call fs_Activa_Dep(False)
   
   pnl_TraDep.Visible = True
   pnl_TraInd.Visible = False
   pnl_Comerc.Visible = False
   pnl_Accion.Visible = False
   pnl_Rentis.Visible = False
   
   If modatecli_g_int_ActEcoCyg = 1 Then
      cmd_Grabar.Enabled = False
      cmd_BorAct.Enabled = False
      cmd_EdiAct.Enabled = False
   ElseIf modatecli_g_int_ActEcoCyg = 2 Then    'Si ya hay datos ingresados
      For r_int_Contad = 1 To UBound(modatecli_g_arr_CygActEco)
         grd_Listad.Rows = grd_Listad.Rows + 1
         
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         Select Case modatecli_g_arr_CygActEco(r_int_Contad).ActEco_OrdAct
            Case 1:  grd_Listad.Text = "PRINCIPAL"
            Case 2:  grd_Listad.Text = "SECUNDARIA"
            Case 9:  grd_Listad.Text = "ANTERIOR"
         End Select
         
         grd_Listad.Col = 2
         grd_Listad.Text = modatecli_g_arr_CygActEco(r_int_Contad).ActEco_OrdAct
         
         grd_Listad.Col = 1
         Select Case modatecli_g_arr_CygActEco(r_int_Contad).ActEco_ActEco
            Case 11: grd_Listad.Text = "EMPLEADO PUBLICO"
            Case 12: grd_Listad.Text = "EMPLEADO PRIVADO"
            Case 21: grd_Listad.Text = "PROFESIONAL INDEPENDIENTE"
            Case 31: grd_Listad.Text = "COMERCIANTE"
            Case 41: grd_Listad.Text = "ACCIONISTA"
            Case 51: grd_Listad.Text = "RENTISTA"
         End Select
      Next r_int_Contad
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 3215
   grd_Listad.ColWidth(1) = 5930
   grd_Listad.ColWidth(2) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   Call modsis_gs_Carga_TipDocIde(cmb_TipDoc, 2)
   Call modsis_gs_Carga_ActEco(cmb_ActEco)
   Call modsis_gs_Carga_OrdAct(cmb_OrdAct)
   Call moddat_gs_Carga_CdCIIU(cmb_CodCiu)
   
   Call gs_Carga_LisIte(cmb_GirCom, l_arr_GirCom, 1, "502")

   Call modsis_gs_Carga_DirTipVia(cmb_TipVia)
   Call modsis_gs_Carga_DirTipZon(cmb_TipZon)
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   
   Call gs_Carga_LisIte(cmb_Dep_NomCar, l_arr_Dep_NomCar, 1, "503")
   Call modsis_gs_Carga_FreHab(cmb_Dep_FreHab)
   
   Call modsis_gs_Carga_AfiNeg(cmb_Ind_ConLoc)
   Call modsis_gs_Carga_TipDocIde(cmb_Ind_TipDoc, 2)
   Call modsis_gs_Carga_RegTri(cmb_Com_RegTri)
   Call modsis_gs_Carga_TipViv_ResAct(cmb_Com_TipLoc, 3)
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

Private Sub cmb_DptDir_Change()
   l_str_DptDir = cmb_DptDir.Text
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

Private Sub fs_Limpia()
   cmb_ActEco.ListIndex = -1
   cmb_OrdAct.ListIndex = -1

   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   pnl_FlgEmp.Visible = False
   cmb_CodCiu.ListIndex = -1
   txt_RazSoc.Text = ""
   txt_Sucurs.Text = ""
   cmb_GirCom.ListIndex = -1
   txt_GirCom.Text = ""
   
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
   txt_Telef1.Text = ""
   txt_NumFax.Text = ""
   
   Call fs_Limpia_TraDep
   Call fs_Limpia_TraInd
   Call fs_Limpia_TraCom
   Call fs_Limpia_TraAcc
End Sub

Private Sub fs_Limpia_TraDep()
   ipp_Dep_IngNet.Value = 0
   cmb_Dep_FreHab.ListIndex = -1
   cmb_Dep_NomCar.ListIndex = -1
   txt_Dep_NomCar.Text = ""
   txt_Dep_NomAre.Text = ""
   ipp_Dep_FecIng.Text = Format(Date - CDate(365), "dd/mm/yyyy")
   txt_Dep_Telefo.Text = ""
   txt_Dep_NumAnx.Text = ""
   txt_Dep_Celula.Text = ""
   txt_Dep_DirEle.Text = ""
   chk_Dep_DirEle.Value = 0
   chk_Dep_DirEle.Enabled = False
   txt_Dep_TlfRhh.Text = ""
   txt_Dep_AnxRhh.Text = ""
   ipp_Dep_FecCes.Text = Format(Date - CDate(365 * 2), "dd/mm/yyyy")
End Sub

Private Sub fs_Limpia_TraInd()
   ipp_Ind_IngNet.Value = 0
   cmb_Ind_ConLoc.ListIndex = -1
   cmb_Ind_TipDoc.ListIndex = -1
   txt_Ind_NumDoc.Text = ""
   txt_Ind_RazSoc.Text = ""
   txt_Ind_Telef1.Text = ""
   txt_Ind_Telef2.Text = ""
   
   cmb_Ind_TipDoc.Enabled = False
   txt_Ind_NumDoc.Enabled = False
   txt_Ind_RazSoc.Enabled = False
   txt_Ind_Telef1.Enabled = False
   txt_Ind_Telef2.Enabled = False
End Sub

Private Sub fs_Limpia_TraCom()
   ipp_Com_IngNet.Value = 0
   ipp_Com_VtaMen.Value = 0
   ipp_Com_PorPar.Value = 0
   ipp_Com_FecIni.Text = Format(Date, "dd/mm/yyyy")
   cmb_Com_RegTri.ListIndex = -1
   
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

Private Sub fs_Limpia_TraAcc()
   ipp_Acc_IngNet.Value = 0
   ipp_Acc_UtiVec.Value = 0
   ipp_Acc_UtiAno.Value = 0
   ipp_Acc_Porcen.Value = 0
End Sub

Private Sub fs_Limpia_TraRen()
   ipp_Ren_IngNet.Value = 0
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_ActEco.Enabled = p_Activa
   cmb_OrdAct.Enabled = p_Activa
   cmb_TipDoc.Enabled = p_Activa
   txt_NumDoc.Enabled = p_Activa
   cmb_CodCiu.Enabled = p_Activa
   txt_RazSoc.Enabled = p_Activa
   txt_Sucurs.Enabled = p_Activa
   cmb_GirCom.Enabled = p_Activa
   txt_GirCom.Enabled = p_Activa
   cmb_TipVia.Enabled = p_Activa
   txt_NomVia.Enabled = p_Activa
   txt_Numero.Enabled = p_Activa
   txt_Interi.Enabled = p_Activa
   cmb_TipZon.Enabled = p_Activa
   txt_NomZon.Enabled = p_Activa
   txt_Refere.Enabled = p_Activa
   cmb_DptDir.Enabled = p_Activa
   cmb_PrvDir.Enabled = p_Activa
   cmb_DstDir.Enabled = p_Activa
   txt_Telefo.Enabled = p_Activa
   txt_Telef1.Enabled = p_Activa
   txt_NumFax.Enabled = p_Activa

   cmd_Agrega.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_NueAct.Enabled = Not p_Activa
   cmd_BorAct.Enabled = Not p_Activa
   cmd_EdiAct.Enabled = Not p_Activa
End Sub

Private Sub fs_Activa_Dep(ByVal p_Activa As Integer)
   ipp_Dep_IngNet.Enabled = p_Activa
   cmb_Dep_FreHab.Enabled = p_Activa
   cmb_Dep_NomCar.Enabled = p_Activa
   txt_Dep_NomCar.Enabled = p_Activa
   txt_Dep_NomAre.Enabled = p_Activa
   ipp_Dep_FecIng.Enabled = p_Activa
   txt_Dep_Telefo.Enabled = p_Activa
   txt_Dep_NumAnx.Enabled = p_Activa
   txt_Dep_Celula.Enabled = p_Activa
   txt_Dep_DirEle.Enabled = p_Activa
   chk_Dep_DirEle.Enabled = p_Activa
   txt_Dep_TlfRhh.Enabled = p_Activa
   txt_Dep_AnxRhh.Enabled = p_Activa
   
   If cmb_OrdAct.ListIndex > -1 Then
      If cmb_OrdAct.ItemData(cmb_OrdAct.ListIndex) = 9 Then
         ipp_Dep_FecCes.Enabled = p_Activa
      End If
   Else
      ipp_Dep_FecCes.Enabled = False
   End If
End Sub

Private Sub fs_Activa_Ind(ByVal p_Activa As Integer)
   ipp_Ind_IngNet.Enabled = p_Activa
   cmb_Ind_ConLoc.Enabled = p_Activa
   cmb_Ind_TipDoc.Enabled = p_Activa
   txt_Ind_NumDoc.Enabled = p_Activa
   txt_Ind_RazSoc.Enabled = p_Activa
   txt_Ind_Telef1.Enabled = p_Activa
   txt_Ind_Telef2.Enabled = p_Activa
End Sub

Private Sub fs_Activa_Com(ByVal p_Activa As Integer)
   ipp_Com_IngNet.Enabled = p_Activa
   ipp_Com_VtaMen.Enabled = p_Activa
   ipp_Com_FecIni.Enabled = p_Activa
   cmb_Com_RegTri.Enabled = p_Activa
   cmb_Com_TipLoc.Enabled = p_Activa
   ipp_Com_AlqMen.Enabled = p_Activa
   ipp_Com_PorPar.Enabled = p_Activa
   txt_Com_NomArr.Enabled = p_Activa
   txt_Com_Tl1Arr.Enabled = p_Activa
   txt_Com_Tl2Arr.Enabled = p_Activa
   txt_Com_NomArr.Enabled = p_Activa
End Sub

Private Sub fs_Activa_Acc(ByVal p_Activa As Integer)
   ipp_Acc_IngNet.Enabled = p_Activa
   ipp_Acc_UtiVec.Enabled = p_Activa
   ipp_Acc_UtiAno.Enabled = p_Activa
   ipp_Acc_Porcen.Enabled = p_Activa
End Sub

Private Sub fs_Activa_Ren(ByVal p_Activa As Integer)
   ipp_Ren_IngNet.Enabled = p_Activa
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_Acc_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Acc_UtiVec)
   End If
End Sub

Private Sub ipp_Acc_Porcen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Agrega)
   End If
End Sub

Private Sub ipp_Acc_UtiAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Acc_Porcen)
   End If
End Sub

Private Sub ipp_Acc_UtiVec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Acc_UtiAno)
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
      Call gs_SetFocus(txt_Dep_Telefo)
   End If
End Sub

Private Sub ipp_Dep_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Dep_FreHab)
   End If
End Sub

Private Sub ipp_Ind_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Ind_ConLoc)
   End If
End Sub

Private Sub ipp_Ren_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Agrega)
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
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-()")
   End If
End Sub

Private Sub txt_Dep_DirEle_Change()
   If Len(Trim(txt_Dep_DirEle)) > 0 Then
      chk_Dep_DirEle.Enabled = True
   Else
      chk_Dep_DirEle.Value = 0
      chk_Dep_DirEle.Enabled = False
   End If
End Sub

Private Sub txt_Dep_DirEle_GotFocus()
   Call gs_SelecTodo(txt_Dep_DirEle)
End Sub

Private Sub txt_Dep_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_TlfRhh)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
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

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_CodCiu.Enabled Then
         Call gs_SetFocus(cmb_CodCiu)
      Else
         Call gs_SetFocus(txt_Sucurs)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumDoc_LostFocus()
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      txt_NumDoc.Text = "00000000000"
   Else
      txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000000")
   End If
   
   If cmb_ActEco.ListIndex = -1 Then
      Exit Sub
   End If
      
   If cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 21 Or cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 51 Then
      Exit Sub
   End If
         
   Call fs_BusEmp
   
   If cmb_CodCiu.Enabled Then
      Call gs_SetFocus(cmb_CodCiu)
   Else
      Call gs_SetFocus(txt_Sucurs)
   End If
End Sub

Private Sub txt_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_RazSoc)
End Sub

Private Sub txt_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Sucurs)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_Sucurs_GotFocus()
   Call gs_SelecTodo(txt_Sucurs)
End Sub

Private Sub txt_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_GirCom.Enabled Then
         Call gs_SetFocus(cmb_GirCom)
      Else
         Call gs_SetFocus(cmb_TipVia)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_GirCom_GotFocus()
   Call gs_SelecTodo(txt_GirCom)
End Sub

Private Sub txt_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
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
      Call gs_SetFocus(txt_Numero)
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

Private Sub cmb_TipZon_Click()
   Call gs_SetFocus(txt_NomZon)
End Sub

Private Sub cmb_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_TipZon_Click
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
      Call gs_SetFocus(txt_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Telef1_GotFocus()
   Call gs_SelecTodo(txt_Telef1)
End Sub

Private Sub txt_Telef1_KeyPress(KeyAscii As Integer)
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
      If pnl_TraDep.Visible Then
         Call gs_SetFocus(ipp_Dep_IngNet)
      ElseIf pnl_TraInd.Visible Then
         Call gs_SetFocus(ipp_Ind_IngNet)
      ElseIf pnl_Comerc.Visible Then
         Call gs_SetFocus(ipp_Com_IngNet)
      ElseIf pnl_Accion.Visible Then
         Call gs_SetFocus(ipp_Acc_IngNet)
      ElseIf pnl_Rentis.Visible Then
         Call gs_SetFocus(ipp_Ren_IngNet)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ind_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_Ind_NumDoc)
End Sub

Private Sub txt_Ind_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ind_RazSoc)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Ind_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_Ind_RazSoc)
End Sub

Private Sub txt_Ind_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ind_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_Ind_Telef1_GotFocus()
   Call gs_SelecTodo(txt_Ind_Telef1)
End Sub

Private Sub txt_Ind_Telef1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Ind_Telef2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_()")
   End If
End Sub

Private Sub txt_Ind_Telef2_GotFocus()
   Call gs_SelecTodo(txt_Ind_Telef2)
End Sub

Private Sub txt_Ind_Telef2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Agrega)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_()")
   End If
End Sub

Private Sub txt_Dep_Telefo_GotFocus()
   Call gs_SelecTodo(txt_Dep_Telefo)
End Sub

Private Sub txt_Dep_Telefo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_NumAnx)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Dep_NumAnx_GotFocus()
   Call gs_SelecTodo(txt_Dep_NumAnx)
End Sub

Private Sub txt_Dep_NumAnx_KeyPress(KeyAscii As Integer)
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
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-()")
   End If
End Sub

Private Sub txt_Dep_TlfRhh_GotFocus()
   Call gs_SelecTodo(txt_Dep_TlfRhh)
End Sub

Private Sub txt_Dep_TlfRhh_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Dep_AnxRhh)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Dep_AnxRhh_GotFocus()
   Call gs_SelecTodo(txt_Dep_AnxRhh)
End Sub

Private Sub txt_Dep_AnxRhh_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_Dep_FecCes.Enabled Then
         Call gs_SetFocus(ipp_Dep_FecCes)
      Else
         Call gs_SetFocus(cmd_Agrega)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Function ff_Valida_TraDep() As Integer
   ff_Valida_TraDep = False
   
   If ipp_Dep_IngNet.Value = 0 Then
      MsgBox "Ingrese el Ingreso Neto.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Dep_IngNet)
      Exit Function
   End If
      
   If cmb_Dep_FreHab.ListIndex = -1 Then
      MsgBox "Seleccione la Frecuencia de Haberes.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Dep_FreHab)
      Exit Function
   End If
   
   If cmb_Dep_NomCar.ListIndex = -1 Then
      MsgBox "Seleccione el Cargo Actual del Cliente.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Dep_NomCar)
      Exit Function
   End If
   
   If Len(Trim(txt_Dep_NomAre.Text)) = 0 Then
      MsgBox "Ingrese el Area para el cual trabaja el cliente.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_Dep_NomAre)
      Exit Function
   End If
   
   If CDate(ipp_Dep_FecIng.Text) > Date Then
      MsgBox "La Fecha de Ingreso no debe ser mayor a la actual.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Dep_FecIng)
      Exit Function
   End If
   
   If Len(Trim(txt_Dep_Telefo.Text)) = 0 Then
      MsgBox "Ingrese el Teléfono y/o Anexo Laboral del Cliente.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_Dep_Telefo)
      Exit Function
   End If
   
   If Len(Trim(txt_Dep_TlfRhh.Text)) = 0 Then
      MsgBox "Ingrese el Teléfono y/o Anexo del Dpto. de Recursos Humanos.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_Dep_TlfRhh)
      Exit Function
   End If
   
   'Si Es Actividad Anterior
   If ipp_Dep_FecCes.Enabled Then
      If CDate(ipp_Dep_FecIng.Text) >= CDate(ipp_Dep_FecCes) Then
         MsgBox "La Fecha de Cese no puede ser menor a la Fecha de Ingreso.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(ipp_Dep_FecCes)
         Exit Function
      End If
   
      If CDate(ipp_Dep_FecCes.Text) > Date Then
         MsgBox "La Fecha de Cese no debe ser mayor a la actual.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(ipp_Dep_FecCes)
         Exit Function
      End If
   End If

   ff_Valida_TraDep = True
End Function

Public Function ff_Valida_TraInd() As Integer
   ff_Valida_TraInd = False
   
   If ipp_Ind_IngNet.Value = 0 Then
      MsgBox "Ingrese el Ingreso Neto Mensual.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Ind_IngNet)
      Exit Function
   End If
   
   If cmb_Ind_ConLoc.ListIndex = -1 Then
      MsgBox "Seleccione si tiene Contrato de Locación.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Ind_ConLoc)
      Exit Function
   End If
   
   'Si tiene Contrato de Locación de Servicios
   If cmb_Ind_ConLoc.ItemData(cmb_Ind_ConLoc.ListIndex) = 1 Then
      If cmb_Ind_TipDoc.ListIndex = -1 Then
         MsgBox "Seleccione si tiene Contrato de Locación.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmb_Ind_ConLoc)
         Exit Function
      End If
      
      If cmb_Ind_TipDoc.ListIndex = -1 Then
         MsgBox "Seleccione el Tipo de Documento de Identidad.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmb_Ind_TipDoc)
         Exit Function
      End If
      
      If Len(Trim(txt_Ind_NumDoc.Text)) <> 11 Then
         MsgBox "Ingrese correctamemte el Número de Documento.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(txt_Ind_NumDoc)
         Exit Function
      End If
      
      If Not gf_Valida_RUC(Mid(txt_Ind_NumDoc.Text, 1, Len(txt_Ind_NumDoc.Text) - 1), Right(txt_Ind_NumDoc.Text, 1)) Then
         MsgBox "Ingrese correctamente el Número de Documento de Identidad.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(txt_Ind_NumDoc)
         Exit Function
      End If
      
      If Len(Trim(txt_Ind_RazSoc.Text)) = 0 Then
         MsgBox "Ingrese la Razón Social.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(txt_Ind_RazSoc)
         Exit Function
      End If
   End If
   
   ff_Valida_TraInd = True
End Function

Private Function ff_Valida_Comerc() As Integer
   ff_Valida_Comerc = False
   
   If ipp_Com_IngNet.Value = 0 Then
      MsgBox "Ingrese el Ingreso Neto Mensual.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Com_IngNet)
      Exit Function
   End If

   If ipp_Com_VtaMen.Value = 0 Then
      MsgBox "Ingrese las Ventas Mensuales.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Com_VtaMen)
      Exit Function
   End If

   If CDate(ipp_Com_FecIni.Text) > Date Then
      MsgBox "Ingrese correctamente la Fecha de Inicio de Operaciones.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Com_FecIni)
      Exit Function
   End If

   If cmb_Com_RegTri.ListIndex = -1 Then
      MsgBox "Seleccione el Régimen Tributario.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Com_RegTri)
      Exit Function
   End If
   
   If ipp_Com_PorPar.Value = 0 Then
      MsgBox "Ingrese el Porcentaje de Participación.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Com_PorPar)
      Exit Function
   End If
   
   If cmb_Com_TipLoc.ListIndex = -1 Then
      MsgBox "Seleccione el Tipo de Local Comercial.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Com_TipLoc)
      Exit Function
   End If
   
   If cmb_Com_TipLoc.ItemData(cmb_Com_TipLoc.ListIndex) = 2 Then
      If ipp_Com_AlqMen.Value = 0 Then
         MsgBox "Ingrese el Alquiler Mensual del Local.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(ipp_Com_AlqMen)
         Exit Function
      End If
   
      If Len(Trim(txt_Com_NomArr.Text)) = 0 Then
         MsgBox "Ingrese el Nombre del Arrendador del Local Comercial.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(txt_Com_NomArr)
         Exit Function
      End If
   
      If Len(Trim(txt_Com_Tl1Arr.Text)) = 0 Then
         MsgBox "Ingrese el Teléfono de Arrendador del Local Comercial.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(txt_Com_Tl1Arr)
         Exit Function
      End If
   End If
   
   ff_Valida_Comerc = True
End Function

Private Function ff_Valida_Accion() As Integer
   ff_Valida_Accion = False
   
   If ipp_Acc_IngNet.Value = 0 Then
      MsgBox "Ingrese el Ingreso Neto Mensual.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Acc_IngNet)
      Exit Function
   End If

   If ipp_Acc_UtiVec.Value = 0 Then
      MsgBox "Ingrese la cantidad de veces que se reparten Utilidades.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Acc_UtiVec)
      Exit Function
   End If

   If ipp_Acc_UtiAno.Value = 0 Then
      MsgBox "Ingrese cada cuantos años se reparten Utilidades.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Acc_UtiAno)
      Exit Function
   End If

   If ipp_Acc_Porcen.Value = 0 Then
      MsgBox "Ingrese el Porcentaje de Accionariado.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Acc_Porcen)
      Exit Function
   End If
   
   ff_Valida_Accion = True
End Function

Private Function ff_Valida_Rentis() As Integer
   ff_Valida_Rentis = False
   
   If ipp_Ren_IngNet.Value = 0 Then
      MsgBox "Ingrese el Ingreso Neto Mensual.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Ren_IngNet)
      Exit Function
   End If
   
   ff_Valida_Rentis = True
End Function

Private Sub fs_Arreglo_Genera(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEco_ActEco = cmb_ActEco.ItemData(cmb_ActEco.ListIndex)
   modatecli_g_arr_CygActEco(p_Indice).ActEco_OrdAct = cmb_OrdAct.ItemData(cmb_OrdAct.ListIndex)
   modatecli_g_arr_CygActEco(p_Indice).ActEco_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   modatecli_g_arr_CygActEco(p_Indice).ActEco_NumDoc = txt_NumDoc.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEco_CodCiu = cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)
   
   If pnl_FlgEmp.Visible And pnl_FlgEmp.Caption = "NR" And (cmb_ActEco.ItemData(cmb_ActEco.ListIndex) <> 21) Then
      modatecli_g_arr_CygActEco(p_Indice).ActEco_EmpReg = 2
   Else
      modatecli_g_arr_CygActEco(p_Indice).ActEco_EmpReg = 1
   End If
   
   modatecli_g_arr_CygActEco(p_Indice).ActEco_RazSoc = txt_RazSoc.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Sucurs = txt_Sucurs.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEco_GiroCd = l_arr_GirCom(cmb_GirCom.ListIndex + 1).Genera_Codigo
   modatecli_g_arr_CygActEco(p_Indice).ActEco_GiroNm = txt_GirCom.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEco_TipVia = cmb_TipVia.ItemData(cmb_TipVia.ListIndex)
   modatecli_g_arr_CygActEco(p_Indice).ActEco_NomVia = txt_NomVia.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Numero = txt_Numero.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Interi = txt_Interi.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEco_TipZon = cmb_TipZon.ItemData(cmb_TipZon.ListIndex)
   modatecli_g_arr_CygActEco(p_Indice).ActEco_NomZon = txt_NomZon.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEco_DptDir = cmb_DptDir.ItemData(cmb_DptDir.ListIndex)
   modatecli_g_arr_CygActEco(p_Indice).ActEco_PrvDir = cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex)
   modatecli_g_arr_CygActEco(p_Indice).ActEco_DstDir = cmb_DstDir.ItemData(cmb_DstDir.ListIndex)
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Refere = txt_Refere.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Telefo = txt_Telefo.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Telef1 = txt_Telef1.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEco_NumFax = txt_NumFax.Text
End Sub

Private Sub fs_Arreglo_TraDep(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_IngNet = ipp_Dep_IngNet.Value
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_FreHab = cmb_Dep_FreHab.ItemData(cmb_Dep_FreHab.ListIndex)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_CargoC = l_arr_Dep_NomCar(cmb_Dep_NomCar.ListIndex + 1).Genera_Codigo
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_CargoN = txt_Dep_NomCar.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_NomAre = txt_Dep_NomAre.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_FecIng = Format(CDate(ipp_Dep_FecIng.Text), "dd/mm/yyyy")
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_Telefo = txt_Dep_Telefo.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_NAnexo = txt_Dep_NumAnx.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_Celula = txt_Dep_Celula.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_DirEle = txt_Dep_DirEle.Text
   
   If chk_Dep_DirEle.Value = 1 Then
      modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_Autori = 1
   Else
      modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_Autori = 2
   End If
   
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_TlfRhh = txt_Dep_TlfRhh.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_AnxRhh = txt_Dep_AnxRhh.Text
   
   If ipp_Dep_FecCes.Enabled Then
      modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_FecCes = Format(CDate(ipp_Dep_FecCes), "dd/mm/yyyy")
   End If
End Sub

Private Sub fs_Arreglo_TraInd(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_IngNet = ipp_Ind_IngNet.Value
   modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_ConLoc = cmb_Ind_ConLoc.ItemData(cmb_Ind_ConLoc.ListIndex)
   
   If cmb_Ind_TipDoc.Enabled Then
      modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_TipDoc = cmb_Ind_TipDoc.ItemData(cmb_Ind_TipDoc.ListIndex)
      modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_NumDoc = txt_Ind_NumDoc.Text
      modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_RazSoc = txt_Ind_RazSoc.Text
      modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_Telef1 = txt_Ind_Telef1.Text
      modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_Telef2 = txt_Ind_Telef2.Text
   End If
End Sub

Private Sub fs_Arreglo_Comerc(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_IngNet = ipp_Com_IngNet.Value
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_VtaMen = ipp_Com_VtaMen.Value
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_FecIni = Format(CDate(ipp_Com_FecIni.Text), "dd/mm/yyyy")
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_RegTri = cmb_Com_RegTri.ItemData(cmb_Com_RegTri.ListIndex)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_TipLoc = cmb_Com_TipLoc.ItemData(cmb_Com_TipLoc.ListIndex)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_AlqMen = ipp_Com_AlqMen.Value
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_PorPar = ipp_Com_PorPar.Value
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_NomArr = txt_Com_NomArr.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_Tl1Arr = txt_Com_Tl1Arr.Text
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_Tl2Arr = txt_Com_Tl2Arr.Text
End Sub

Private Sub fs_Arreglo_Accion(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoAcc_IngNet = ipp_Acc_IngNet.Value
   modatecli_g_arr_CygActEco(p_Indice).ActEcoAcc_UtiVec = ipp_Acc_UtiVec.Value
   modatecli_g_arr_CygActEco(p_Indice).ActEcoAcc_UtiAno = ipp_Acc_UtiAno.Value
   modatecli_g_arr_CygActEco(p_Indice).ActEcoAcc_Porcen = ipp_Acc_Porcen.Value
End Sub

Private Sub fs_Arreglo_Rentis(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoRen_IngNet = ipp_Ren_IngNet.Value
End Sub

Private Sub fs_Limpia_Arreglo(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEco_ActEco = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEco_OrdAct = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEco_EmpReg = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEco_TipDoc = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEco_NumDoc = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_CodCiu = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEco_RazSoc = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Sucurs = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_GiroCd = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_GiroNm = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_TipVia = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEco_NomVia = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Numero = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Interi = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_TipZon = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEco_NomZon = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_DptDir = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEco_PrvDir = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEco_DstDir = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Refere = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Telefo = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_Telef1 = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEco_NumFax = ""
End Sub

Private Sub fs_Limpia_Arreglo_TraDep(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_IngNet = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_FreHab = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_CargoC = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_CargoN = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_NomAre = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_FecIng = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_Telefo = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_NAnexo = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_Celula = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_DirEle = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_Autori = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_TlfRhh = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_AnxRhh = ""
   
   If ipp_Dep_FecCes.Enabled Then
      modatecli_g_arr_CygActEco(p_Indice).ActEcoDep_FecCes = ""
   End If
End Sub

Private Sub fs_Limpia_Arreglo_TraInd(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_IngNet = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_ConLoc = 0
   
   If cmb_Ind_TipDoc.Enabled Then
      modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_TipDoc = 0
      modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_NumDoc = ""
      modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_RazSoc = ""
      modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_Telef1 = ""
      modatecli_g_arr_CygActEco(p_Indice).ActEcoInd_Telef2 = ""
   End If
End Sub

Private Sub fs_Limpia_Arreglo_Comerc(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_IngNet = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_VtaMen = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_FecIni = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_RegTri = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_PorPar = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_TipLoc = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_AlqMen = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_NomArr = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_Tl1Arr = ""
   modatecli_g_arr_CygActEco(p_Indice).ActEcoCom_Tl2Arr = ""
End Sub

Private Sub fs_Limpia_Arreglo_Accion(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoAcc_IngNet = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoAcc_UtiVec = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoAcc_UtiAno = 0
   modatecli_g_arr_CygActEco(p_Indice).ActEcoAcc_Porcen = 0
End Sub

Private Sub fs_Limpia_Arreglo_Rentis(ByVal p_Indice As Integer)
   modatecli_g_arr_CygActEco(p_Indice).ActEcoRen_IngNet = 0
End Sub

Private Sub fs_BorIte(ByVal p_Item As Integer)
   Dim r_int_Contad     As Integer
   
   If UBound(modatecli_g_arr_CygActEco) = 1 Then
      ReDim modatecli_g_arr_CygActEco(0)
      Exit Sub
   End If
   
   For r_int_Contad = p_Item To UBound(modatecli_g_arr_CygActEco) - 1
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_ActEco = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_ActEco
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_OrdAct = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_OrdAct
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_EmpReg = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_EmpReg
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_TipDoc = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_TipDoc
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_NumDoc = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_NumDoc
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_CodCiu = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_CodCiu
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_RazSoc = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_RazSoc
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_Sucurs = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_Sucurs
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_GiroCd = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_GiroCd
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_GiroNm = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_GiroNm
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_TipVia = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_TipVia
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_NomVia = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_NomVia
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_Numero = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_Numero
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_Interi = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_Interi
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_TipZon = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_TipZon
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_NomZon = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_NomZon
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_DptDir = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_DptDir
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_PrvDir = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_PrvDir
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_DstDir = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_DstDir
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_Refere = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_Refere
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_Telefo = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_Telefo
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_Telef1 = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_Telef1
      modatecli_g_arr_CygActEco(r_int_Contad).ActEco_NumFax = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEco_NumFax
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_IngNet = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_IngNet
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_FreHab = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_FreHab
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_CargoC = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_CargoC
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_CargoN = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_CargoN
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_NomAre = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_NomAre
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_FecIng = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_FecIng
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_Telefo = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_Telefo
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_NAnexo = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_NAnexo
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_Celula = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_Celula
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_TlfRhh = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_TlfRhh
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_AnxRhh = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_AnxRhh
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_DirEle = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_DirEle
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_Autori = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_Autori
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoDep_FecCes = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoDep_FecCes
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoInd_IngNet = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoInd_IngNet
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoInd_ConLoc = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoInd_ConLoc
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoInd_TipDoc = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoInd_TipDoc
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoInd_NumDoc = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoInd_NumDoc
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoInd_RazSoc = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoInd_RazSoc
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoInd_Telef1 = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoInd_Telef1
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoInd_Telef2 = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoInd_Telef2
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoCom_IngNet = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoCom_IngNet
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoCom_VtaMen = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoCom_VtaMen
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoCom_PorPar = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoCom_PorPar
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoCom_FecIni = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoCom_FecIni
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoCom_RegTri = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoCom_RegTri
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoCom_TipLoc = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoCom_TipLoc
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoCom_AlqMen = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoCom_AlqMen
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoCom_NomArr = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoCom_NomArr
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoCom_Tl1Arr = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoCom_Tl1Arr
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoCom_Tl2Arr = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoCom_Tl2Arr
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoAcc_IngNet = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoAcc_IngNet
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoAcc_UtiVec = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoAcc_UtiVec
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoAcc_UtiAno = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoAcc_UtiAno
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoAcc_Porcen = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoAcc_Porcen
      modatecli_g_arr_CygActEco(r_int_Contad).ActEcoRen_IngNet = modatecli_g_arr_CygActEco(r_int_Contad + 1).ActEcoRen_IngNet
   Next r_int_Contad
   
   ReDim Preserve modatecli_g_arr_CygActEco(UBound(modatecli_g_arr_CygActEco) - 1)
End Sub

Private Sub fs_BusEmp()
   If cmb_TipDoc.ListIndex > -1 Then
      'Obteniendo Información de la Empresa
      g_str_Parame = "SP_EMP$DATGEN_CONSULTA " & Format(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), "0") & ", " & txt_NumDoc.Text
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
         Exit Sub
      End If
      
      pnl_FlgEmp.Visible = True
      
      'Si Empresa está registrada
      If g_rst_Princi.RecordCount > 0 Then
         g_rst_Princi.MoveFirst
         
         Select Case g_rst_Princi!DATGEN_CLASIF
            Case 1: pnl_FlgEmp.Caption = "A"
            Case 2: pnl_FlgEmp.Caption = "B"
            Case 3: pnl_FlgEmp.Caption = "C"
            Case 8: pnl_FlgEmp.Caption = "NN"
            Case 9: pnl_FlgEmp.Caption = "PC"
         End Select
         
         'Cargando Información de la Empresa
         Call gs_BuscarCombo_Item(cmb_CodCiu, Trim(g_rst_Princi!DATGEN_COCIIU))
         txt_RazSoc.Text = Trim(g_rst_Princi!DATGEN_RAZSOC)
   
         cmb_GirCom.ListIndex = gf_Busca_Arregl(l_arr_GirCom, Trim(g_rst_Princi!DATGEN_GCOMCO)) - 1
         txt_GirCom.Text = Trim(g_rst_Princi!DATGEN_GCOMNO)
         
         Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!DATGEN_TIPVIA)
         txt_NomVia.Text = Trim(g_rst_Princi!DATGEN_NOMVIA)
         txt_Numero.Text = Trim(g_rst_Princi!DATGEN_NUMERO)
         txt_Interi.Text = Trim(g_rst_Princi!DATGEN_INTDPT)
         Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!DATGEN_TIPZON)
         txt_NomZon.Text = Trim(g_rst_Princi!DATGEN_NOMZON)
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(Trim(g_rst_Princi!DatGen_UBIGEO), 2)))
         If cmb_PrvDir.ListCount = 0 Then
            Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         End If
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(Trim(g_rst_Princi!DatGen_UBIGEO), 3, 2)))
         
         If cmb_DstDir.ListCount = 0 Then
            Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         End If
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(Trim(g_rst_Princi!DatGen_UBIGEO), 2)))
         txt_Refere.Text = Trim(g_rst_Princi!DatGen_Refere)
         txt_Telefo.Text = Trim(g_rst_Princi!DATGEN_TELEFO)
         txt_Telef1.Text = Trim(g_rst_Princi!DATGEN_TELEF1)
         txt_NumFax.Text = Trim(g_rst_Princi!DatGen_NUMFAX)
         
         cmb_CodCiu.Enabled = False
         cmb_GirCom.Enabled = False
         txt_GirCom.Enabled = False
         txt_RazSoc.Enabled = False
      Else
         pnl_FlgEmp.Caption = "NR"
      
         cmb_CodCiu.Enabled = True
         cmb_GirCom.Enabled = True
         txt_GirCom.Enabled = True
         txt_RazSoc.Enabled = True
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Else
      pnl_FlgEmp.Visible = False
   End If
End Sub

