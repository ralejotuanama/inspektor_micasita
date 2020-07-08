VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   1650
   ClientTop       =   1335
   ClientWidth     =   11670
   Icon            =   "AteCli_frm_109.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7755
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
      _ExtentY        =   13679
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
         TabIndex        =   28
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
            TabIndex        =   29
            Top             =   60
            Width           =   7305
            _Version        =   65536
            _ExtentX        =   12885
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes - Datos del Cónyuge"
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
            Picture         =   "AteCli_frm_109.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   4875
         Left            =   30
         TabIndex        =   30
         Top             =   2040
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   8599
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
         Begin VB.TextBox txt_NDoAlt 
            Height          =   315
            Left            =   8220
            MaxLength       =   12
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TDoAlt 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DocAlt 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   60
            Width           =   1065
         End
         Begin VB.ComboBox cmb_ActEco 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   4500
            Width           =   765
         End
         Begin VB.TextBox txt_ApeCas 
            Height          =   315
            Left            =   8190
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1200
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1200
            Width           =   3315
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   870
            Width           =   3315
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1530
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Paises 
            Height          =   315
            Left            =   2010
            TabIndex        =   11
            Text            =   "cmb_Paises"
            Top             =   2190
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptNac 
            Height          =   315
            Left            =   8190
            TabIndex        =   12
            Text            =   "cmb_DptNac"
            Top             =   2190
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvNac 
            Height          =   315
            Left            =   2010
            TabIndex        =   13
            Text            =   "cmb_PrvNac"
            Top             =   2520
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstNac 
            Height          =   315
            Left            =   8190
            TabIndex        =   14
            Text            =   "cmb_DstNac"
            Top             =   2520
            Width           =   3315
         End
         Begin VB.ComboBox cmb_NivEst 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2850
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Profes 
            Height          =   315
            Left            =   8190
            TabIndex        =   16
            Text            =   "cmb_Profes"
            Top             =   2850
            Width           =   3315
         End
         Begin VB.TextBox txt_Celula 
            Height          =   315
            Left            =   2010
            MaxLength       =   9
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   3180
            Width           =   3315
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   8190
            MaxLength       =   120
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   3180
            Width           =   1665
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
            Left            =   9930
            TabIndex        =   19
            Top             =   3210
            Width           =   1485
         End
         Begin VB.ComboBox cmb_ClaSbs 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   3690
            Width           =   3315
         End
         Begin VB.ComboBox cmb_ClasMC 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   4020
            Width           =   3315
         End
         Begin EditLib.fpDateTime ipp_FecNac 
            Height          =   315
            Left            =   2010
            TabIndex        =   10
            Top             =   1860
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   90
            Left            =   60
            TabIndex        =   31
            Top             =   3540
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   90
            Left            =   60
            TabIndex        =   54
            Top             =   4380
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   90
            Left            =   60
            TabIndex        =   56
            Top             =   750
            Width           =   11460
            _Version        =   65536
            _ExtentX        =   20214
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
         Begin VB.Label Label33 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   6240
            TabIndex        =   59
            Top             =   390
            Width           =   1065
         End
         Begin VB.Label Label34 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   58
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label Label35 
            Caption         =   "Personal FF.AA / FF.PP:"
            Height          =   315
            Left            =   90
            TabIndex        =   57
            Top             =   60
            Width           =   1845
         End
         Begin VB.Label Label6 
            Caption         =   "Registra Activ. Econ.:"
            Height          =   285
            Left            =   90
            TabIndex        =   55
            Top             =   4530
            Width           =   1785
         End
         Begin VB.Label Label29 
            Caption         =   "Apellido Casada:"
            Height          =   285
            Left            =   6210
            TabIndex        =   46
            Top             =   1200
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   90
            TabIndex        =   45
            Top             =   1200
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   90
            TabIndex        =   44
            Top             =   870
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   90
            TabIndex        =   43
            Top             =   1530
            Width           =   1485
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha de Nacimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   42
            Top             =   1860
            Width           =   1905
         End
         Begin VB.Label Label8 
            Caption         =   "Nacionalidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   41
            Top             =   2190
            Width           =   1905
         End
         Begin VB.Label Label9 
            Caption         =   "Dpto. Nacimiento:"
            Height          =   315
            Left            =   6210
            TabIndex        =   40
            Top             =   2190
            Width           =   1905
         End
         Begin VB.Label Label10 
            Caption         =   "Provincia Nacimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   39
            Top             =   2520
            Width           =   1905
         End
         Begin VB.Label Label11 
            Caption         =   "Distrito Nacimiento:"
            Height          =   315
            Left            =   6210
            TabIndex        =   38
            Top             =   2520
            Width           =   1905
         End
         Begin VB.Label Label14 
            Caption         =   "Nivel de Estudio:"
            Height          =   315
            Left            =   90
            TabIndex        =   37
            Top             =   2850
            Width           =   1905
         End
         Begin VB.Label Label15 
            Caption         =   "Profesión o Actividad:"
            Height          =   315
            Left            =   6210
            TabIndex        =   36
            Top             =   2850
            Width           =   1905
         End
         Begin VB.Label Label16 
            Caption         =   "Teléfono Celular:"
            Height          =   285
            Left            =   90
            TabIndex        =   35
            Top             =   3180
            Width           =   1485
         End
         Begin VB.Label Label17 
            Caption         =   "E-mail:"
            Height          =   285
            Left            =   6210
            TabIndex        =   34
            Top             =   3180
            Width           =   1485
         End
         Begin VB.Label Label30 
            Caption         =   "Clasificación SBS:"
            Height          =   315
            Left            =   90
            TabIndex        =   33
            Top             =   3690
            Width           =   1545
         End
         Begin VB.Label Label31 
            Caption         =   "Clasificación miCasita:"
            Height          =   315
            Left            =   90
            TabIndex        =   32
            Top             =   4020
            Width           =   1785
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   735
         Left            =   30
         TabIndex        =   47
         Top             =   6960
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
            Picture         =   "AteCli_frm_109.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_ActEco 
            Height          =   675
            Left            =   720
            Picture         =   "AteCli_frm_109.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Actividades Económicas"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10860
            Picture         =   "AteCli_frm_109.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Left            =   30
         TabIndex        =   48
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
            TabIndex        =   50
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   51
         Top             =   1230
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   10170
            Picture         =   "AteCli_frm_109.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   9480
            Picture         =   "AteCli_frm_109.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10860
            Picture         =   "AteCli_frm_109.frx":1248
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   53
            Top             =   60
            Width           =   1845
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Doc. Identidad:"
            Height          =   285
            Left            =   90
            TabIndex        =   52
            Top             =   390
            Width           =   1635
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_09"
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
Dim l_int_FlgCmb     As Integer

Private Sub cmb_ActEco_Click()
   If cmb_ActEco.ListIndex > -1 Then
      If cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 1 Then
         cmd_ActEco.Enabled = True
         
         Call gs_SetFocus(cmd_ActEco)
      Else
         cmd_ActEco.Enabled = False
         Call gs_SetFocus(cmd_Grabar)
      End If
   Else
      cmd_ActEco.Enabled = False
   End If
End Sub

Private Sub cmb_ActEco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ActEco_Click
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
   Call gs_SetFocus(cmb_ActEco)
End Sub

Private Sub cmb_ClasMC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ClasMC_Click
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
         
            Call gs_SetFocus(cmb_NivEst)
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
   
      Call gs_SetFocus(cmb_NivEst)
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
         
            Call gs_SetFocus(cmb_NivEst)
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
   
         Call gs_SetFocus(cmb_NivEst)
      End If
   End If
End Sub

Private Sub cmb_Paises_LostFocus()
   Call SendMessage(cmb_Paises.hWnd, CB_SHOWDROPDOWN, 0, 0&)
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
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
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
         Call gs_SetFocus(cmb_NivEst)
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
      
      Call gs_SetFocus(cmb_NivEst)
   End If
End Sub

Private Sub cmb_DstNac_LostFocus()
   Call SendMessage(cmb_DstNac.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_PrvNac_LostFocus()
   Call SendMessage(cmb_PrvNac.hWnd, CB_SHOWDROPDOWN, 0, 0&)
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

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
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

Private Sub cmd_ActEco_Click()
   moddat_g_str_CygNom = txt_ApePat.Text & " " & txt_ApeMat.Text & " " & txt_Nombre
   moddat_g_int_TipCli = 2
   
   frm_MntCli_03.Show 1
End Sub

Private Sub cmd_Buscar_Click()
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

   moddat_g_int_CygTDo = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   moddat_g_str_CygNDo = txt_NumDoc.Text
      
   If Len(Trim(moddat_g_str_CodPrd)) > 0 Then
      'Verificando que Cónyuge no haya sido ingresado como Cliente Negativo
      If Not atecli_gf_Buscar_BasNeg(moddat_g_int_CygTDo, moddat_g_str_CygNDo) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      'Validar que Cónyuge no tenga una Solicitud de Crédito en Evaluación Como Titular
      If Not atecli_gf_Buscar_SolVig(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      'Validar que Cónyuge no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
      If Not atecli_gf_Buscar_SolVig(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      'Buscando Operaciones de Crédito
      Call atecli_gs_Buscar_CreHip(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
      
      If UBound(modatecli_g_arr_CygOpe) > 0 Then
         MsgBox "El Cónyuge ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
         Call cmd_Limpia_Click
         
         Exit Sub
      End If
   End If
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_CygTDo) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_CygNDo & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      moddat_g_int_FlgCyg = 2
      
      Call gs_BuscarCombo_Item(cmb_DocAlt, g_rst_Princi!DatGen_FLGDOA)
      
      Call gs_BuscarCombo_Item(cmb_TDoAlt, g_rst_Princi!DatGen_TIPDOA)
      txt_NDoAlt.Text = Trim(g_rst_Princi!DatGen_NUMDOA & "")
      
      txt_ApePat.Text = Trim(g_rst_Princi!DatGen_ApePat & "")
      txt_ApeMat.Text = Trim(g_rst_Princi!DatGen_ApeMat & "")
      txt_ApeCas.Text = Trim(g_rst_Princi!DatGen_ApeCas & "")
      txt_Nombre.Text = Trim(g_rst_Princi!DatGen_Nombre & "")
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
         
      Call gs_BuscarCombo_Item(cmb_NivEst, g_rst_Princi!DatGen_NivEst)
      cmb_Profes.ListIndex = gf_Busca_Arregl(l_arr_Profes, g_rst_Princi!DatGen_Profes) - 1
         
      txt_Celula.Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      txt_DirEle.Text = Trim(g_rst_Princi!DatGen_DirEle & "")
      
      If g_rst_Princi!DATGEN_AUTENV = 1 Then
         chk_DirEle.Value = 1
         chk_DirEle.Enabled = True
      End If
         
      Call gs_BuscarCombo_Text(cmb_ClaSbs, g_rst_Princi!DATGEN_CLASBS, 1)
      Call gs_BuscarCombo_Text(cmb_ClasMC, g_rst_Princi!DATGEN_CLASMC, 1)
      
      Call gs_BuscarCombo_Item(cmb_ActEco, g_rst_Princi!DATGEN_ACTECO)
      
      If cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 1 Then
         cmd_ActEco.Enabled = True
      End If
      
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      cmd_Buscar.Enabled = False
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   cmb_DocAlt.Enabled = True
   txt_ApePat.Enabled = True
   txt_ApeMat.Enabled = True
   txt_ApeCas.Enabled = True
   txt_Nombre.Enabled = True
   ipp_FecNac.Enabled = True
   cmb_Paises.Enabled = True
   cmb_DptNac.Enabled = True
   cmb_NivEst.Enabled = True
   txt_DirEle.Enabled = True
   txt_Celula.Enabled = True
   cmb_ClaSbs.Enabled = True
   cmb_ClasMC.Enabled = True
   cmb_ActEco.Enabled = True
   
   Call gs_SetFocus(cmb_DocAlt)
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_EdaMin     As Integer
   Dim r_int_EdaMax     As Integer
   Dim r_int_EdaAct     As Integer
   
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
   
   
   If cmb_DocAlt.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Cliente es miembro de las FF.AA o FF.PP.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DocAlt)
      Exit Sub
   End If
   
   If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
      If cmb_TDoAlt.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TDoAlt)
         Exit Sub
      End If
      
      If Len(Trim(txt_NDoAlt.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NDoAlt)
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
   
   If frm_MntCli_02.cmb_CodSex.ItemData(frm_MntCli_02.cmb_CodSex.ListIndex) = 2 Then
      If Len(Trim(txt_ApeCas.Text)) > 0 Then
         MsgBox "El Cónyuge no puede presentar Apellido de Casada.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_ApeCas)
         Exit Sub
      End If
   End If
   
   If Not IsDate(ipp_FecNac.Text) Then
      MsgBox "La Fecha de Nacimiento no es válida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecNac)
      Exit Sub
   End If
   
   If CDate(ipp_FecNac.Text) > Date Then
      MsgBox "Debe ingresar una Fecha de Nacimiento valida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecNac)
      Exit Sub
   End If
   
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
         Exit Sub
      End If
   End If
   
   If cmb_Paises.ListIndex = -1 Then
      MsgBox "Debe seleccionar el País de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Paises)
      Exit Sub
   End If
   
   If CInt(l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo) = 4028 Then
      If cmb_DptNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Departamento de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DptNac)
         Exit Sub
      End If
      
      If cmb_PrvNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Provincia de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PrvNac)
         Exit Sub
      End If
      
      If cmb_DstNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Distrito de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DstNac)
         Exit Sub
      End If
   End If
   
   If cmb_NivEst.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Nivel de Estudio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NivEst)
      Exit Sub
   End If
   
   If cmb_Profes.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Profesión u Oficio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Profes)
      Exit Sub
   End If
   
   If cmb_ClaSbs.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Clasificación de la SBS.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ClaSbs)
      Exit Sub
   End If
   
   If cmb_ClasMC.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Clasificación en miCasita.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ClasMC)
      Exit Sub
   End If
   
   If cmb_ActEco.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Cliente registra Actividad Económica.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ActEco)
      Exit Sub
   End If
   
   If cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 1 Then
      If moddat_g_arr_ActEco_Cyg(1).ActEco_TipAct = 0 Then
         MsgBox "Debe registrar las Actividades Económicas del Cliente.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_ActEco)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_Inicia_DatCyg
   
   moddat_g_arr_DatCyg(1).DatCli_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   moddat_g_arr_DatCyg(1).DatCli_NumDoc = txt_NumDoc.Text
   moddat_g_arr_DatCyg(1).DatCli_DocAlt = cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex)
   
   If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
      moddat_g_arr_DatCyg(1).DatCli_TDoAlt = cmb_TDoAlt.ItemData(cmb_TDoAlt.ListIndex)
      moddat_g_arr_DatCyg(1).DatCli_NDoAlt = txt_NDoAlt.Text
   End If
   
   moddat_g_arr_DatCyg(1).DatCli_ApePat = txt_ApePat.Text
   moddat_g_arr_DatCyg(1).DatCli_ApeMat = txt_ApeMat.Text
   moddat_g_arr_DatCyg(1).DatCli_ApeCas = txt_ApeCas.Text
   moddat_g_arr_DatCyg(1).DatCli_Nombre = txt_Nombre.Text
   moddat_g_arr_DatCyg(1).DatCli_FecNac = ipp_FecNac.Text
   moddat_g_arr_DatCyg(1).DatCli_Paises = l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo
   
   If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo = "004028" Then
      moddat_g_arr_DatCyg(1).DatCli_UbiGeo = Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00") & Format(cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex), "00") & Format(cmb_DstNac.ItemData(cmb_DstNac.ListIndex), "00")
   Else
      moddat_g_arr_DatCyg(1).DatCli_UbiGeo = "000000"
   End If
   
   moddat_g_arr_DatCyg(1).DatCli_NivEst = cmb_NivEst.ItemData(cmb_NivEst.ListIndex)
   moddat_g_arr_DatCyg(1).DatCli_Profes = l_arr_Profes(cmb_Profes.ListIndex + 1).Genera_Codigo
   moddat_g_arr_DatCyg(1).DatCli_Celula = txt_Celula.Text
   moddat_g_arr_DatCyg(1).DatCli_DirEle = txt_DirEle.Text
   
   If chk_DirEle.Value = 1 Then
      moddat_g_arr_DatCyg(1).DatCli_ChkEle = 1
   ElseIf chk_DirEle.Value = 0 Then
      moddat_g_arr_DatCyg(1).DatCli_ChkEle = 2
   End If
   
   moddat_g_arr_DatCyg(1).DatCli_ClaSbs = Left(cmb_ClaSbs.Text, 1)
   moddat_g_arr_DatCyg(1).DatCli_ClasMC = Left(cmb_ClasMC.Text, 1)
   
   moddat_g_arr_DatCyg(1).DatCli_ActEco = cmb_ActEco.ItemData(cmb_ActEco.ListIndex)
   
   If cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 2 Then
      Call moddat_gs_Inicia_ActEco(2, 1)
      Call moddat_gs_Inicia_ActEco(2, 2)
   End If
   
   moddat_g_str_CygNom = Trim(txt_ApePat.Text) & " " & Trim(txt_ApeMat.Text) & " " & Trim(txt_Nombre.Text)
   
   moddat_g_int_FlgCyg = 2
   
   Unload Me
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipDoc)
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
   
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Limpia
   
   If moddat_g_int_FlgCyg = 2 Then
      'Leer información del Arreglo
      Call gs_BuscarCombo_Item(cmb_TipDoc, moddat_g_arr_DatCyg(1).DatCli_TipDoc)
      txt_NumDoc.Text = moddat_g_arr_DatCyg(1).DatCli_NumDoc
      
      Call gs_BuscarCombo_Item(cmb_DocAlt, moddat_g_arr_DatCyg(1).DatCli_DocAlt)
      
      If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
         Call gs_BuscarCombo_Item(cmb_TDoAlt, moddat_g_arr_DatCyg(1).DatCli_TDoAlt)
         txt_NDoAlt.Text = moddat_g_arr_DatCyg(1).DatCli_NDoAlt
      End If
      
      txt_ApePat.Text = moddat_g_arr_DatCyg(1).DatCli_ApePat
      txt_ApeMat.Text = moddat_g_arr_DatCyg(1).DatCli_ApeMat
      txt_ApeCas.Text = moddat_g_arr_DatCyg(1).DatCli_ApeCas
      txt_Nombre.Text = moddat_g_arr_DatCyg(1).DatCli_Nombre
      
      ipp_FecNac.Text = moddat_g_arr_DatCyg(1).DatCli_FecNac
      
      cmb_Paises.ListIndex = gf_Busca_Arregl(l_arr_Paises, moddat_g_arr_DatCyg(1).DatCli_Paises) - 1
      
      If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo = "004028" Then
         Call gs_BuscarCombo_Item(cmb_DptNac, CInt(Left(moddat_g_arr_DatCyg(1).DatCli_UbiGeo, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvNac, Left(moddat_g_arr_DatCyg(1).DatCli_UbiGeo, 2))
         Call gs_BuscarCombo_Item(cmb_PrvNac, CInt(Mid(moddat_g_arr_DatCyg(1).DatCli_UbiGeo, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstNac, Left(moddat_g_arr_DatCyg(1).DatCli_UbiGeo, 2), Mid(moddat_g_arr_DatCyg(1).DatCli_UbiGeo, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstNac, CInt(Right(moddat_g_arr_DatCyg(1).DatCli_UbiGeo, 2)))
      End If
      
      Call gs_BuscarCombo_Item(cmb_NivEst, moddat_g_arr_DatCyg(1).DatCli_NivEst)
      cmb_Profes.ListIndex = gf_Busca_Arregl(l_arr_Profes, moddat_g_arr_DatCyg(1).DatCli_Profes) - 1
      
      txt_Celula.Text = moddat_g_arr_DatCyg(1).DatCli_Celula
      txt_DirEle.Text = moddat_g_arr_DatCyg(1).DatCli_DirEle
      
      If moddat_g_arr_DatCyg(1).DatCli_ChkEle = 1 Then
         chk_DirEle.Value = 1
      Else
         chk_DirEle.Value = 0
      End If
      
      Call gs_BuscarCombo_Text(cmb_ClaSbs, moddat_g_arr_DatCyg(1).DatCli_ClaSbs, 1)
      Call gs_BuscarCombo_Text(cmb_ClasMC, moddat_g_arr_DatCyg(1).DatCli_ClasMC, 1)
   
      Call gs_BuscarCombo_Item(cmb_ActEco, moddat_g_arr_DatCyg(1).DatCli_ActEco)
      
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      cmd_Buscar.Enabled = False
      
      cmb_DocAlt.Enabled = True
      txt_ApePat.Enabled = True
      txt_ApeMat.Enabled = True
      txt_ApeCas.Enabled = True
      txt_Nombre.Enabled = True
      ipp_FecNac.Enabled = True
      cmb_Paises.Enabled = True
      cmb_DptNac.Enabled = True
      cmb_NivEst.Enabled = True
      txt_DirEle.Enabled = True
      txt_Celula.Enabled = True
      cmb_ClaSbs.Enabled = True
      cmb_ClasMC.Enabled = True
      cmb_ActEco.Enabled = True
      
      If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
         cmb_TDoAlt.Enabled = True
         txt_NDoAlt.Enabled = True
      End If
      
      If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo = "004028" Then
         cmb_DptNac.Enabled = True
         cmb_PrvNac.Enabled = True
         cmb_DstNac.Enabled = True
      End If
      
      Call gs_SetFocus(cmb_DocAlt)
   Else
      Call gs_SetFocus(cmb_TipDoc)
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   Call moddat_gs_Carga_LisIte_Combo(cmb_DocAlt, 1, "214")

   Call moddat_gs_Carga_LisIte_Combo(cmb_ActEco, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TDoAlt, 1, "231")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_NivEst, 1, "209")
   
   Call moddat_gs_Carga_LisIte(cmb_Paises, l_arr_Paises, 1, "500")
   Call moddat_gs_Carga_LisIte(cmb_Profes, l_arr_Profes, 1, "501")
      
   Call moddat_gs_Carga_Depart(cmb_DptNac)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_ClaSbs, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ClasMC, 1, "058")
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""

   cmb_TipDoc.Enabled = True
   txt_NumDoc.Enabled = True
   cmd_Buscar.Enabled = True

   cmb_DocAlt.ListIndex = -1
   cmb_TDoAlt.ListIndex = -1
   txt_NDoAlt.Text = ""
   
   cmb_TDoAlt.Enabled = False
   txt_NDoAlt.Enabled = False
   
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_ApeCas.Text = ""
   txt_Nombre.Text = ""
   
   ipp_FecNac.Text = Format(Date, "dd/mm/yyyy")
   cmb_Paises.ListIndex = -1
   cmb_DptNac.ListIndex = -1
   cmb_PrvNac.Clear
   cmb_DstNac.Clear
   cmb_DptNac.Enabled = False
   cmb_PrvNac.Enabled = False
   cmb_DstNac.Enabled = False
   cmb_NivEst.ListIndex = -1
   cmb_Profes.ListIndex = -1
   txt_DirEle.Text = ""
   txt_Celula.Text = ""
   
   chk_DirEle.Value = 0
   chk_DirEle.Enabled = False
   
   cmb_ClaSbs.ListIndex = -1
   cmb_ClasMC.ListIndex = -1

   cmb_DocAlt.Enabled = False
   cmb_TDoAlt.Enabled = False
   txt_NDoAlt.Enabled = False
   txt_ApePat.Enabled = False
   txt_ApeMat.Enabled = False
   txt_ApeCas.Enabled = False
   txt_Nombre.Enabled = False
   ipp_FecNac.Enabled = False
   cmb_Paises.Enabled = False
   cmb_DptNac.Enabled = False
   cmb_PrvNac.Enabled = False
   cmb_DstNac.Enabled = False
   cmb_NivEst.Enabled = False
   txt_DirEle.Enabled = False
   chk_DirEle.Enabled = False
   txt_Celula.Enabled = False
   cmb_ClaSbs.Enabled = False
   cmb_ClasMC.Enabled = False
   cmb_ActEco.Enabled = False
   cmd_ActEco.Enabled = False
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_TipDoc.Enabled = p_Activa
   txt_NumDoc.Enabled = p_Activa
   
   cmb_DocAlt.Enabled = Not p_Activa
   cmb_TDoAlt.Enabled = Not p_Activa
   txt_NDoAlt.Enabled = Not p_Activa
   
   txt_ApePat.Enabled = Not p_Activa
   txt_ApeMat.Enabled = Not p_Activa
   txt_ApeCas.Enabled = Not p_Activa
   txt_Nombre.Enabled = Not p_Activa
   ipp_FecNac.Enabled = Not p_Activa
   cmb_Paises.Enabled = Not p_Activa
   cmb_DptNac.Enabled = Not p_Activa
   cmb_PrvNac.Enabled = Not p_Activa
   cmb_DstNac.Enabled = Not p_Activa
   cmb_NivEst.Enabled = Not p_Activa
   cmb_Profes.Enabled = Not p_Activa
   txt_Celula.Enabled = Not p_Activa
   txt_DirEle.Enabled = Not p_Activa
   chk_DirEle.Enabled = Not p_Activa
   cmb_ClaSbs.Enabled = Not p_Activa
   cmb_ClasMC.Enabled = Not p_Activa
   cmb_ActEco.Enabled = Not p_Activa
End Sub

Private Sub ipp_FecNac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Paises)
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
      Call gs_SetFocus(cmb_ClaSbs)
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
      Call gs_SetFocus(ipp_FecNac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
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


