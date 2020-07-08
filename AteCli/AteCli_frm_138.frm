VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_SegSol_16 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   11070
   ClientLeft      =   -525
   ClientTop       =   570
   ClientWidth     =   15420
   Icon            =   "AteCli_frm_138.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11070
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16125
      _Version        =   65536
      _ExtentX        =   28443
      _ExtentY        =   19500
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
         Height          =   735
         Left            =   30
         TabIndex        =   1
         Top             =   10260
         Width           =   16005
         _Version        =   65536
         _ExtentX        =   28231
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
         Begin VB.CommandButton cmd_SegSol 
            Height          =   675
            Left            =   15300
            Picture         =   "AteCli_frm_138.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Seguimiento de Solicitud"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   3
         Top             =   750
         Width           =   16005
         _Version        =   65536
         _ExtentX        =   28231
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   15300
            Picture         =   "AteCli_frm_138.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   14610
            Picture         =   "AteCli_frm_138.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   675
            Left            =   13920
            Picture         =   "AteCli_frm_138.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar Cliente"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   13230
            Picture         =   "AteCli_frm_138.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   60
            Width           =   4905
         End
         Begin VB.CheckBox chk_Produc 
            Caption         =   "Todos los Productos"
            Height          =   315
            Left            =   1410
            TabIndex        =   4
            Top             =   390
            Width           =   2685
         End
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1410
            TabIndex        =   21
            Top             =   1050
            Width           =   1485
            _Version        =   196608
            _ExtentX        =   2619
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
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1410
            TabIndex        =   22
            Top             =   720
            Width           =   1485
            _Version        =   196608
            _ExtentX        =   2619
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
         Begin VB.Label Label2 
            Caption         =   "Fecha Fin:"
            Height          =   315
            Left            =   90
            TabIndex        =   24
            Top             =   1050
            Width           =   1155
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   90
            TabIndex        =   23
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   10
            Top             =   60
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   16005
         _Version        =   65536
         _ExtentX        =   28231
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
            TabIndex        =   12
            Top             =   60
            Width           =   8355
            _Version        =   65536
            _ExtentX        =   14737
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitudes Desembolsadas de Crédito Hipotecario"
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
            Picture         =   "AteCli_frm_138.frx":1076
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   7995
         Left            =   30
         TabIndex        =   13
         Top             =   2220
         Width           =   16005
         _Version        =   65536
         _ExtentX        =   28231
         _ExtentY        =   14102
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   7605
            Left            =   60
            TabIndex        =   14
            Top             =   360
            Width           =   15885
            _ExtentX        =   28019
            _ExtentY        =   13414
            _Version        =   393216
            Rows            =   45
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   4980
            TabIndex        =   15
            Top             =   60
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "DOI Cliente"
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   3450
            TabIndex        =   16
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
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
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   6450
            TabIndex        =   17
            Top             =   60
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
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
         Begin Threed.SSPanel pnl_Tit_FecDes 
            Height          =   285
            Left            =   10260
            TabIndex        =   18
            Top             =   60
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Desembolso"
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   90
            TabIndex        =   19
            Top             =   60
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
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
         Begin Threed.SSPanel pnl_Tit_ConHip 
            Height          =   285
            Left            =   11520
            TabIndex        =   20
            Top             =   60
            Width           =   2520
            _Version        =   65536
            _ExtentX        =   4445
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Consejero Hipotecario"
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
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   14040
            TabIndex        =   25
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
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
      End
   End
End
Attribute VB_Name = "frm_SegSol_16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera

Private Sub chk_Produc_Click()
   If chk_Produc.Value = 1 Then
      cmb_Produc.ListIndex = -1
      cmb_Produc.Enabled = False
      
      Call gs_SetFocus(ipp_FecIni)
   ElseIf chk_Produc.Value = 0 Then
      cmb_Produc.Enabled = True
      
      Call gs_SetFocus(cmb_Produc)
   End If
End Sub

Private Sub cmb_Produc_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If chk_Produc.Value = 0 Then
      If cmb_Produc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Produc)
         Exit Sub
      End If
   End If
   
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   Call fs_Buscar
End Sub

Private Sub cmd_BusCli_Click()
   frm_SegSol_25.Show 1
End Sub

Private Sub cmd_Limpia_Click()
   cmb_Produc.ListIndex = -1
   chk_Produc.Value = 0
   
   ipp_FecIni.Text = Format(Date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(Date, "dd/mm/yyyy")
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_Produc)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_Produc.Enabled = p_Activa
   chk_Produc.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   ipp_FecIni.Enabled = p_Activa
   ipp_FecFin.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_SegSol.Enabled = Not p_Activa
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SegSol_Click()
   grd_Listad.Col = 1
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
         
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_SegSol_26.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Activa(True)
   
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_Produc(cmb_Produc, l_arr_Produc, 4)
   
   grd_Listad.ColWidth(0) = 3365
   grd_Listad.ColWidth(1) = 1535
   grd_Listad.ColWidth(2) = 1475
   grd_Listad.ColWidth(3) = 3815
   grd_Listad.ColWidth(4) = 1265
   grd_Listad.ColWidth(5) = 2510
   grd_Listad.ColWidth(6) = 1535
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   Call cmd_SegSol_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Buscar()
   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, CRE_HIPMAE B WHERE "
   
   If chk_Produc.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   'Si Tipo de Usuario es Consejero Hipotecario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = HIPMAE_NUMSOL AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 2 AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO ASC"
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Solicitudes para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      
      grd_Listad.Col = 4
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CONHIP)
      'grd_Listad.Text = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_CONHIP))
      
      grd_Listad.Col = 6
      grd_Listad.Text = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
      
      grd_Listad.Col = 7
      grd_Listad.Text = g_rst_Princi!SOLMAE_CODPRD

      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_FECDES)
      
      g_rst_Princi.MoveNext
   Loop
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad)

   Screen.MousePointer = 0
   
   Call fs_Activa(False)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub pnl_Tit_ConHip_Click()
   If Len(Trim(pnl_Tit_ConHip.Tag)) = 0 Or pnl_Tit_ConHip.Tag = "D" Then
      pnl_Tit_ConHip.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_ConHip.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecDes_Click()
   If Len(Trim(pnl_Tit_FecDes.Tag)) = 0 Or pnl_Tit_FecDes.Tag = "D" Then
      pnl_Tit_FecDes.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "N")
   Else
      pnl_Tit_FecDes.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "N-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumSol_Click()
   If Len(Trim(pnl_Tit_NumSol.Tag)) = 0 Or pnl_Tit_NumSol.Tag = "D" Then
      pnl_Tit_NumSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NumSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub
