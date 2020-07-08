VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_IngSol_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form7"
   ClientHeight    =   5190
   ClientLeft      =   2070
   ClientTop       =   3075
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5175
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   9128
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   735
         Left            =   30
         TabIndex        =   39
         Top             =   4380
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_Operac 
            Height          =   675
            Left            =   750
            Picture         =   "AteCli_frm_008.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Lista de Operaciones Crediticias"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_LisRec 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_008.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Lista de Solicitudes Rechazadas"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_008.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_ActEco 
            Height          =   675
            Left            =   9450
            Picture         =   "AteCli_frm_008.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Actividades Económicas"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Acepta 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_008.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Aceptar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1138
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   495
            Left            =   630
            TabIndex        =   38
            Top             =   60
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Datos del Cónyuge"
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
            Left            =   3690
            TabIndex        =   23
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
            Height          =   540
            Left            =   60
            Picture         =   "AteCli_frm_008.frx":106A
            Top             =   60
            Width           =   495
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2745
         Left            =   30
         TabIndex        =   24
         Top             =   1590
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   4842
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
         Begin VB.TextBox txt_ApeCas 
            Height          =   315
            Left            =   8130
            MaxLength       =   30
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   390
            Width           =   3345
         End
         Begin VB.ComboBox cmb_ActEco 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2370
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   2130
            MaxLength       =   30
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   2130
            MaxLength       =   30
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   2130
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Paises 
            Height          =   315
            Left            =   2130
            TabIndex        =   8
            Text            =   "cmb_Paises"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptNac 
            Height          =   315
            Left            =   8130
            TabIndex        =   9
            Text            =   "cmb_DptNac"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvNac 
            Height          =   315
            Left            =   2130
            TabIndex        =   10
            Text            =   "cmb_PrvNac"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstNac 
            Height          =   315
            Left            =   8130
            TabIndex        =   11
            Text            =   "cmb_DstNac"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_NivEst 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Profes 
            Height          =   315
            Left            =   8130
            TabIndex        =   13
            Text            =   "cmb_Profes"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.TextBox txt_Celula 
            Height          =   315
            Left            =   2130
            MaxLength       =   12
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   2040
            Width           =   3315
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   8130
            MaxLength       =   120
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   2040
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
            Left            =   9870
            TabIndex        =   16
            Top             =   2070
            Width           =   1485
         End
         Begin EditLib.fpDateTime ipp_FecNac 
            Height          =   315
            Left            =   8130
            TabIndex        =   7
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
         Begin Threed.SSPanel pnl_EdaCli 
            Height          =   315
            Left            =   9480
            TabIndex        =   42
            Top             =   720
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "240 "
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
         Begin Threed.SSPanel pnl_MesCli 
            Height          =   315
            Left            =   10440
            TabIndex        =   44
            Top             =   690
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "240 "
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
         Begin VB.Label Label31 
            Caption         =   "Meses"
            Height          =   285
            Left            =   10950
            TabIndex        =   45
            Top             =   750
            Width           =   555
         End
         Begin VB.Label Label30 
            Caption         =   "Años"
            Height          =   285
            Left            =   10020
            TabIndex        =   43
            Top             =   750
            Width           =   555
         End
         Begin VB.Label Label1 
            Caption         =   "Apellido de Casada:"
            Height          =   285
            Left            =   6060
            TabIndex        =   41
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "Registra Activ. Econom.:"
            Height          =   315
            Left            =   90
            TabIndex        =   37
            Top             =   2370
            Width           =   1905
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   60
            TabIndex        =   36
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   90
            TabIndex        =   35
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   90
            TabIndex        =   34
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha de Nacimiento:"
            Height          =   315
            Left            =   6090
            TabIndex        =   33
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label Label8 
            Caption         =   "Nacionalidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   32
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label9 
            Caption         =   "Dpto. Nacimiento:"
            Height          =   315
            Left            =   6090
            TabIndex        =   31
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label10 
            Caption         =   "Provincia Nacimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   30
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label11 
            Caption         =   "Distrito Nacimiento:"
            Height          =   315
            Left            =   6090
            TabIndex        =   29
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label14 
            Caption         =   "Nivel de Estudio:"
            Height          =   315
            Left            =   90
            TabIndex        =   28
            Top             =   1710
            Width           =   1905
         End
         Begin VB.Label Label15 
            Caption         =   "Profesión:"
            Height          =   315
            Left            =   6090
            TabIndex        =   27
            Top             =   1710
            Width           =   1905
         End
         Begin VB.Label Label16 
            Caption         =   "Teléfono Celular:"
            Height          =   285
            Left            =   90
            TabIndex        =   26
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label Label17 
            Caption         =   "E-mail:"
            Height          =   285
            Left            =   6090
            TabIndex        =   25
            Top             =   2040
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   825
         Left            =   30
         TabIndex        =   46
         Top             =   720
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   2130
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   420
            Width           =   2415
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   3315
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   10110
            Picture         =   "AteCli_frm_008.frx":11E1
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_008.frx":14EB
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.Label Label18 
            Caption         =   "Nro. Docum. Identidad:"
            Height          =   285
            Left            =   90
            TabIndex        =   48
            Top             =   390
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   47
            Top             =   90
            Width           =   1845
         End
      End
   End
End
Attribute VB_Name = "frm_IngSol_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Paises()   As moddat_tpo_Genera
Dim l_arr_Profes()   As moddat_tpo_Genera
Dim l_int_FlgCmb     As Integer
Dim l_str_Paises     As String
Dim l_str_Profes     As String
Dim l_str_DptNac     As String
Dim l_str_PrvNac     As String
Dim l_str_DstNac     As String

Private Sub cmb_ActEco_Click()
   cmd_ActEco.Enabled = False
   Call gs_SetFocus(cmd_Acepta)
   
   If cmb_ActEco.ListIndex > -1 Then
      If cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 1 Then
         cmd_ActEco.Enabled = True
         Call gs_SetFocus(cmd_ActEco)
      Else
         ReDim modatecli_g_arr_ActEcoCyg(0)
      End If
   End If
End Sub

Private Sub cmb_ActEco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ActEco_Click
   End If
End Sub

Private Sub cmb_Paises_Change()
   l_str_Paises = cmb_Paises.Text
End Sub

Private Sub cmb_Paises_Click()
   If cmb_Paises.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DptNac.Enabled = True
         cmb_PrvNac.Enabled = True
         cmb_DstNac.Enabled = True
         
         Call gs_SetFocus(cmb_DptNac)
         
         If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo <> "004028" Then
            cmb_DptNac.ListIndex = -1
            cmb_PrvNac.Clear
            cmb_DstNac.Clear
            
            cmb_DptNac.Enabled = False
            cmb_PrvNac.Enabled = False
            cmb_DstNac.Enabled = False
         
            Call gs_SetFocus(cmb_NivEst)
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
   l_int_FlgCmb = True
   l_str_Paises = cmb_Paises.Text
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

      Call gs_SetFocus(cmb_DptNac)
      
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

Private Sub cmd_Acepta_Click()
   If Not ff_Valida() Then
      Exit Sub
   End If
   
   If cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 1 Then
      'Validar que haya ingresado las Actividades Económicas del Cónyuge
      If modatecli_g_int_ActEcoCyg = 1 Then
         MsgBox "Debe registrar las Actividades Económicas del Cónyuge.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_ActEco)
         Exit Sub
      End If
   End If
   
   
   Call modatecli_gs_Limpia_DatGen(2)
   
   'Grabar Datos al Arreglo
   modatecli_g_arr_DatGen(2).DatGen_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   modatecli_g_arr_DatGen(2).DatGen_NumDoc = txt_NumDoc.Text
   modatecli_g_arr_DatGen(2).DatGen_ApePat = txt_ApePat.Text
   modatecli_g_arr_DatGen(2).DatGen_ApeMat = txt_ApeMat.Text
   modatecli_g_arr_DatGen(2).DatGen_ApeCas = txt_ApeCas.Text
   modatecli_g_arr_DatGen(2).DatGen_Nombre = txt_Nombre.Text
   modatecli_g_arr_DatGen(2).DatGen_FecNac = Format(CDate(ipp_FecNac.Text), "dd/mm/yyyy")
   modatecli_g_arr_DatGen(2).DatGen_Paises = l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo
   
   'Titular Hombre Cónyuge Mujer / Tituluar Mujer Cónyuge Hombre
   If frm_IngSol_01.cmb_CodSex.ItemData(frm_IngSol_01.cmb_CodSex.ListIndex) = 1 Then
      modatecli_g_arr_DatGen(2).DatGen_CodSex = 2
   ElseIf frm_IngSol_01.cmb_CodSex.ItemData(frm_IngSol_01.cmb_CodSex.ListIndex) = 2 Then
      modatecli_g_arr_DatGen(2).DatGen_CodSex = 1
   End If
   
   If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo = "004028" Then
      modatecli_g_arr_DatGen(2).DatGen_DptNac = cmb_DptNac.ItemData(cmb_DptNac.ListIndex)
      modatecli_g_arr_DatGen(2).DatGen_PrvNac = cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex)
      modatecli_g_arr_DatGen(2).DatGen_DstNac = cmb_DstNac.ItemData(cmb_DstNac.ListIndex)
   End If
   
   modatecli_g_arr_DatGen(2).DatGen_NivEst = cmb_NivEst.ItemData(cmb_NivEst.ListIndex)
   modatecli_g_arr_DatGen(2).DatGen_Profes = l_arr_Profes(cmb_Profes.ListIndex + 1).Genera_Codigo
   modatecli_g_arr_DatGen(2).DatGen_Celula = txt_Celula.Text
   modatecli_g_arr_DatGen(2).DatGen_DirEle = txt_DirEle.Text
   
   If chk_DirEle.Value = 1 Then
      modatecli_g_arr_DatGen(2).DatGen_Autori = 1
   Else
      modatecli_g_arr_DatGen(2).DatGen_Autori = 2
   End If
   
   modatecli_g_arr_DatGen(2).DatGen_ActEco = cmb_ActEco.ItemData(cmb_ActEco.ListIndex)
   
   If cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 2 Then
      ReDim modatecli_g_arr_Cyg_ActEco(0)
      
      modatecli_g_int_ActPri_Cyg = 0
      modatecli_g_int_ActSec_Cyg = 0
   End If
   
   modatecli_g_int_CygDatGen = 2
   
   Unload Me
End Sub

Private Sub cmd_ActEco_Click()
   If Not ff_Valida() Then
      Exit Sub
   End If
   
   modatecli_g_int_Tip_ActEco = 2
   frm_IngSol_02.Show 1
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_con_PltPar
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      Case 1
         If Len(Trim(txt_NumDoc.Text)) <> 8 Then
            MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_PltPar
            Call gs_SetFocus(txt_NumDoc)
            Exit Sub
         End If

      Case 2
         If Len(Trim(txt_NumDoc.Text)) < 0 Then
            MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_PltPar
            Call gs_SetFocus(txt_NumDoc)
            Exit Sub
         End If
      
      Case 3
         If Len(Trim(txt_NumDoc.Text)) < 0 Then
            MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_PltPar
            Call gs_SetFocus(txt_NumDoc)
            Exit Sub
         End If
   End Select
   
   Call modatecli_gs_Limpia_DatGen(2)     'Cónyuge - Datos Generales
   ReDim modatecli_g_arr_Cyg_ActEco(0)    'Cónyuge - Datos Económicos
   
   'Datos Actividades Económicas Cliente Cónyuge
   modatecli_g_str_CodCiu_Cyg = ""
   modatecli_g_str_GirCom_Cyg = ""
   modatecli_g_str_SecEco_Cyg = ""
   modatecli_g_int_TDoEmp_Cyg = 0
   modatecli_g_str_NDoEmp_Cyg = ""
   modatecli_g_int_ActSec_Cyg = 0
      
   atecli_int_CliCyg = 1              'Flag de Registrado en Base de Datos (1 = No / 2 = Si) (Cónyuge)
   
   'Inicializando Arreglos de Solicitudes Rechazadas
   ReDim modatecli_g_arr_CygRec(0)

   'Inicializando Flag de Datos Ingresados
   modatecli_g_int_CygDatGen = 1
   
   'VALIDACIONES DE CLIENTE
   'Validar que Cliente no se encuentre en Base Negativa
   If Not atecli_gf_Buscar_BasNeg(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If

   'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Titular
   If Not atecli_gf_Buscar_SolVig(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, 1) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
   If Not atecli_gf_Buscar_SolVig(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, 2) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   'Buscando Solicitudes Rechazadas
   Call atecli_gs_Buscar_SolRec(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, 2)
   
   If UBound(modatecli_g_arr_CygRec) > 0 Then
      cmd_LisRec.Visible = True
   End If
   
   'Buscando Operaciones
   Call atecli_gs_Buscar_CreHip(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, 2)
   
   If UBound(modatecli_g_arr_CygOpe) > 0 Then
      If moddat_g_str_CodPrd = "001" Then    'Si Producto es Mivivienda
         MsgBox "El Cliente ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      cmd_Operac.Visible = True
   End If
   
   'Buscando Información de Cliente Cónyuge
   Call atecli_gs_Buscar_DatCli(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, 2)
   
   'Activando Controles
   Call fs_Activa(False)
   
   'Si se encontro Cliente en Base de Datos Asignar Información de Cliente Titular a Controles
   If atecli_int_CliCyg = 2 Then
      Call fs_Arreglo_DatCli
   End If

   Call gs_SetFocus(txt_ApePat)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call modatecli_gs_Limpia_DatGen(2)     'Cónyuge - Datos Generales
   ReDim modatecli_g_arr_Cyg_ActEco(0)    'Cónyuge - Datos Económicos
   
   'Datos Actividades Económicas Cliente Cónyuge
   modatecli_g_str_CodCiu_Cyg = ""
   modatecli_g_str_GirCom_Cyg = ""
   modatecli_g_str_SecEco_Cyg = ""
   modatecli_g_int_TDoEmp_Cyg = 0
   modatecli_g_str_NDoEmp_Cyg = ""
   modatecli_g_int_ActSec_Cyg = 0
      
   atecli_int_CliCyg = 1              'Flag de Registrado en Base de Datos (1 = No / 2 = Si) (Cónyuge)
   
   'Inicializando Arreglos de Solicitudes Rechazadas
   ReDim modatecli_g_arr_CygRec(0)

   'Inicializando Flag de Datos Ingresados
   modatecli_g_int_CygDatGen = 1
   modatecli_g_int_ActEcoCyg = 1

   Call fs_Activa(True)
   Call fs_Limpia
   
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmd_LisRec_Click()
   frm_LisRec_02.Show 1
End Sub

Private Sub cmd_Operac_Click()
   frm_LisOpe_02.Show 1
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("Al salir de esta manera perderá la información ingresada. ¿Está seguro de salir de la ventana?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt & " Ingreso de Solicitud de Crédito"
   
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicio
   Call fs_Limpia
   
   If modatecli_g_int_CygDatGen = 1 Then
      modatecli_g_int_ActEcoCyg = 1
      cmd_ActEco.Enabled = False
      
      Call fs_Activa(True)
   ElseIf modatecli_g_int_CygDatGen = 2 Then
      Call fs_Arreglo_DatCli
   
      If moddat_g_int_CygTDo > 0 Then
         Call fs_Activa(False)
         
         cmd_Buscar.Enabled = False
         cmd_Limpia.Enabled = False
         
         'Buscando Solicitudes Rechazadas
         Call atecli_gs_Buscar_SolRec(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
      
         If UBound(modatecli_g_arr_LisRec) > 0 Then
            cmd_LisRec.Visible = True
         End If
      
         'Buscando Operaciones
         Call atecli_gs_Buscar_CreHip(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text, 2)
         
         If UBound(modatecli_g_arr_CygOpe) > 0 Then
            If moddat_g_str_CodPrd = "001" Then    'Si Producto es Mivivienda
               MsgBox "El Cliente ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
               Call cmd_Limpia_Click
               Exit Sub
            End If
            
            cmd_Operac.Visible = True
         End If
      Else
         Call fs_Activa(True)
      End If
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   
   txt_ApePat.Enabled = Not p_Habilita
   txt_ApeMat.Enabled = Not p_Habilita
   txt_Nombre.Enabled = Not p_Habilita
   ipp_FecNac.Enabled = Not p_Habilita
   cmb_Paises.Enabled = Not p_Habilita
   cmb_DptNac.Enabled = Not p_Habilita
   cmb_PrvNac.Enabled = Not p_Habilita
   cmb_DstNac.Enabled = Not p_Habilita
   cmb_NivEst.Enabled = Not p_Habilita
   cmb_Profes.Enabled = Not p_Habilita
   cmb_ActEco.Enabled = Not p_Habilita
   
   txt_Celula.Enabled = Not p_Habilita
   txt_DirEle.Enabled = Not p_Habilita
   
   chk_DirEle.Enabled = Not p_Habilita
End Sub

Private Sub fs_Arreglo_DatCli()
   l_int_FlgCmb = True
   
   Call gs_BuscarCombo_Item(cmb_TipDoc, modatecli_g_arr_DatGen(2).DatGen_TipDoc)
   
   txt_NumDoc.Text = modatecli_g_arr_DatGen(2).DatGen_NumDoc
   txt_ApePat.Text = modatecli_g_arr_DatGen(2).DatGen_ApePat
   txt_ApeMat.Text = modatecli_g_arr_DatGen(2).DatGen_ApeMat
   txt_ApeCas.Text = modatecli_g_arr_DatGen(2).DatGen_ApeCas
   txt_Nombre.Text = modatecli_g_arr_DatGen(2).DatGen_Nombre
   ipp_FecNac.Text = modatecli_g_arr_DatGen(2).DatGen_FecNac
   cmb_Paises.ListIndex = gf_Busca_Arregl(l_arr_Paises, modatecli_g_arr_DatGen(2).DatGen_Paises) - 1
   
   If modatecli_g_arr_DatGen(2).DatGen_Paises = "004028" Then
      Call gs_BuscarCombo_Item(cmb_DptNac, modatecli_g_arr_DatGen(2).DatGen_DptNac)
      Call gs_BuscarCombo_Item(cmb_PrvNac, modatecli_g_arr_DatGen(2).DatGen_PrvNac)
      Call gs_BuscarCombo_Item(cmb_DstNac, modatecli_g_arr_DatGen(2).DatGen_DstNac)
   End If
   
   Call gs_BuscarCombo_Item(cmb_NivEst, modatecli_g_arr_DatGen(2).DatGen_NivEst)
   cmb_Profes.ListIndex = gf_Busca_Arregl(l_arr_Profes, modatecli_g_arr_DatGen(2).DatGen_Profes) - 1
   Call gs_BuscarCombo_Item(cmb_ActEco, modatecli_g_arr_DatGen(2).DatGen_ActEco)
   
   txt_Celula.Text = modatecli_g_arr_DatGen(2).DatGen_Celula
   txt_DirEle.Text = modatecli_g_arr_DatGen(2).DatGen_DirEle
   
   If modatecli_g_arr_DatGen(2).DatGen_Autori = 1 Then
      chk_DirEle.Value = 1
   End If
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_ApeCas.Text = ""
   txt_Nombre.Text = ""
   
   ipp_FecNac.Text = Format(CDate(moddat_g_str_FecSis) - CDate(18 * 365.25), "dd/mm/yyyy")
   pnl_EdaCli.Caption = "0 "
   pnl_MesCli.Caption = "0 "
   
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
   chk_DirEle.Value = 0
   chk_DirEle.Enabled = False
   txt_Celula.Text = ""
   
   cmb_ActEco.ListIndex = -1

   cmd_LisRec.Visible = False
   cmd_Operac.Visible = False
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)
   Call moddat_gs_Carga_LisIte_Combo(cmb_NivEst, 1, "209")
   
   Call moddat_gs_Carga_LisIte(cmb_Paises, l_arr_Paises, 1, "500")
   Call moddat_gs_Carga_LisIte(cmb_Profes, l_arr_Profes, 1, "501")
      
   Call moddat_gs_Carga_Depart(cmb_DptNac)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_ActEco, 1, "214")
End Sub

Private Sub ipp_FecNac_Change()
   pnl_EdaCli.Caption = Left(gs_CalcularEdad(CDate(ipp_FecNac.Text), Date), 2) & " "
   pnl_MesCli.Caption = Right(gs_CalcularEdad(CDate(ipp_FecNac.Text), Date), 2) & " "
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
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-()")
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

Private Sub cmb_NivEst_Click()
   Call gs_SetFocus(cmb_Profes)
End Sub

Private Sub cmb_NivEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_NivEst_Click
   End If
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

Private Sub txt_DirEle_Change()
   If Len(Trim(txt_DirEle)) > 0 Then
      chk_DirEle.Enabled = True
   Else
      chk_DirEle.Value = 0
      chk_DirEle.Enabled = False
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

Private Sub txt_DirEle_GotFocus()
   Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ActEco)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
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
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Function ff_Valida() As Integer
   ff_Valida = False
   
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Function
   End If
   
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Function
   End If
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
      txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
   End If
   
   If (txt_NumDoc.Text = moddat_g_str_NumDoc) And (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = moddat_g_int_TipDoc) Then
      MsgBox "El Número de Documento de Identidad del Cónyuge es igual al del Titular.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Function
   End If
   
   If Len(Trim(txt_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Function
   End If
   
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Function
   End If
   
   If CDate(ipp_FecNac.Text) > Date Then
      MsgBox "Debe ingresar una Fecha de Nacimiento valida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecNac)
      Exit Function
   End If
   
   Call moddat_gs_FecSis
   
   If Not (CInt(pnl_EdaCli.Caption) >= modatecli_g_int_Par_EdaMin) Then
      MsgBox "El cliente no cumple con los requisitos de Edad requeridos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecNac)
      Exit Function
   End If
   
   If Not ((CInt(pnl_EdaCli.Caption) < modatecli_g_int_Par_EdaMax) Or (CInt(pnl_EdaCli.Caption) = modatecli_g_int_Par_EdaMax And CInt(pnl_MesCli.Caption) = 0)) Then
      MsgBox "El cliente no cumple con los requisitos de Edad requeridos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecNac)
      Exit Function
   End If
   
   If cmb_Paises.ListIndex = -1 Then
      MsgBox "Debe seleccionar el País de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Paises)
      Exit Function
   End If
   
   If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo = "000001" Then
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

   If cmb_ActEco.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Cónyuge registra Actividad Económica.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ActEco)
      Exit Function
   End If
   
   ff_Valida = True
End Function
