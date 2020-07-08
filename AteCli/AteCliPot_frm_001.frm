VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_IngCliPot_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15060
   Icon            =   "AteCliPot_frm_001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7995
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   15060
      _Version        =   65536
      _ExtentX        =   26564
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   13
         Top             =   750
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   620
            Picture         =   "AteCliPot_frm_001.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Modificar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "AteCliPot_frm_001.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14370
            Picture         =   "AteCliPot_frm_001.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1215
            Picture         =   "AteCliPot_frm_001.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
            Left            =   600
            TabIndex        =   15
            Top             =   60
            Width           =   8835
            _Version        =   65536
            _ExtentX        =   15584
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Registro de Clientes Potenciales"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Picture         =   "AteCliPot_frm_001.frx":0D6C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   5265
         Left            =   30
         TabIndex        =   16
         Top             =   2685
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
         _ExtentY        =   9287
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
         Begin Threed.SSPanel pnl_Proyec 
            Height          =   285
            Left            =   6240
            TabIndex        =   21
            Top             =   60
            Width           =   3240
            _Version        =   65536
            _ExtentX        =   5715
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Proyecto"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   4890
            Left            =   60
            TabIndex        =   10
            Top             =   360
            Width           =   14880
            _ExtentX        =   26247
            _ExtentY        =   8625
            _Version        =   393216
            Rows            =   20
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Telefo 
            Height          =   285
            Left            =   9375
            TabIndex        =   17
            Top             =   60
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Teléfono"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Ape_Nom 
            Height          =   285
            Left            =   1320
            TabIndex        =   18
            Top             =   60
            Width           =   4965
            _Version        =   65536
            _ExtentX        =   8758
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_NumDoc 
            Height          =   285
            Left            =   90
            TabIndex        =   19
            Top             =   60
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Numero Doc."
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Situac 
            Height          =   285
            Left            =   11640
            TabIndex        =   20
            Top             =   60
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2417
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_FecReg 
            Height          =   285
            Left            =   10560
            TabIndex        =   22
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1940
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Registro"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Consej 
            Height          =   285
            Left            =   12960
            TabIndex        =   23
            Top             =   60
            Width           =   1680
            _Version        =   65536
            _ExtentX        =   2963
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Consejero"
            ForeColor       =   16777215
            BackColor       =   16384
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   1230
         Left            =   30
         TabIndex        =   24
         Top             =   1440
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
         _ExtentY        =   2170
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
         Begin VB.CheckBox chk_FecCon 
            Caption         =   "Todos los Registros"
            Height          =   315
            Left            =   6570
            TabIndex        =   5
            Top             =   855
            Width           =   2325
         End
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   3525
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   8835
            Picture         =   "AteCliPot_frm_001.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Buscar Datos"
            Top             =   180
            Width           =   585
         End
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   540
            Width           =   3525
         End
         Begin VB.CheckBox chk_ConHip 
            Caption         =   "Todos los Consejeros"
            Height          =   330
            Left            =   1305
            TabIndex        =   2
            Top             =   855
            Width           =   1875
         End
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   330
            Left            =   6570
            TabIndex        =   4
            Top             =   540
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
            _ExtentY        =   582
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
            Height          =   330
            Left            =   6570
            TabIndex        =   3
            Top             =   180
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
            _ExtentY        =   582
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Situacion :"
            Height          =   195
            Left            =   360
            TabIndex        =   28
            Top             =   180
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Consejero :"
            Height          =   195
            Left            =   360
            TabIndex        =   27
            Top             =   540
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio :"
            Height          =   195
            Left            =   5490
            TabIndex        =   26
            Top             =   180
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Fin :"
            Height          =   195
            Left            =   5490
            TabIndex        =   25
            Top             =   540
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frm_IngCliPot_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_ConHip()      As moddat_tpo_Genera

Private Sub chk_ConHip_Click()
   If chk_ConHip.Value = 1 Then
      cmb_ConHip.ListIndex = -1
      cmb_ConHip.Enabled = False
      Call gs_SetFocus(ipp_FecIni)
   ElseIf chk_ConHip.Value = 0 Then
      cmb_ConHip.Enabled = True
      Call gs_SetFocus(cmb_ConHip)
   End If
End Sub

Private Sub Chk_FecCon_Click()
   If chk_FecCon.Value = 1 Then
      ipp_FecIni.Enabled = False
      ipp_FecFin.Enabled = False
      Call gs_SetFocus(cmd_Buscar)
   ElseIf chk_FecCon.Value = 0 Then
      ipp_FecIni.Enabled = True
      ipp_FecFin.Enabled = True
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub cmb_ConHip_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_ConHip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ConHip_Click
   End If
End Sub

Private Sub cmb_Situac_Click()
   Call gs_SetFocus(cmb_ConHip)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Situac_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
    moddat_g_int_FlgGrb = 1
    frm_IngCliPot_02.SSPanel7.Caption = frm_IngCliPot_02.SSPanel7.Caption & " - Insertar"
    frm_IngCliPot_02.Show 1
    
    Screen.MousePointer = 11
    Call cmd_Buscar_Click
    Screen.MousePointer = 0
    
    
   
End Sub

Private Sub cmd_Buscar_Click()
   '* Validaciones
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar La Situacion.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   If (chk_ConHip.Value = 0) And (cmb_ConHip.ListIndex = -1) Then
      MsgBox "Debe seleccionar un Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ConHip)
      Exit Sub
   End If
   
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_ExpExc.Enabled = False
   grd_Listad.Enabled = False
   Call gs_LimpiaGrid(grd_Listad)
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT POSMAE_NUMSOL, POSMAE_CODCON, PROCLI_TIPDOC, PROCLI_NUMDOC, PROCLI_APEPAT, PROCLI_APEMAT, "
   g_str_Parame = g_str_Parame & "       PROCLI_NOMBRE, POSMAE_FECCON, POSMAE_PROYEC, PROCLI_NUMTEL, PROCLI_NUMCEL, PARDES_DESCRI "
   g_str_Parame = g_str_Parame & "  FROM CRE_POSMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PROCLI B ON B.PROCLI_TIPDOC = A.POSMAE_TIPDOC AND B.PROCLI_NUMDOC = A.POSMAE_NUMDOC "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = '300' AND D.PARDES_CODITE = A.POSMAE_SITUAC "
   If modgen_g_int_TipUsu = 20121 Then   'Si Tipo de Usuario es Consejero Hipotecario
      If cmb_Situac.ItemData(cmb_Situac.ListIndex) <> 0 Then
         g_str_Parame = g_str_Parame & " WHERE A.POSMAE_CODCON = '" & modgen_g_str_CodUsu & "' AND POSMAE_FECCON > 0 AND POSMAE_SITUAC = '" & Format(cmb_Situac.ItemData(cmb_Situac.ListIndex), "000000") & "' "
         If chk_FecCon.Value = 0 Then
            g_str_Parame = g_str_Parame & " AND POSMAE_FECCON BETWEEN " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ""
         End If
      Else
         g_str_Parame = g_str_Parame & " WHERE A.POSMAE_CODCON = '" & modgen_g_str_CodUsu & "' AND POSMAE_FECCON > 0 "
         If chk_FecCon.Value = 0 Then
            g_str_Parame = g_str_Parame & " AND POSMAE_FECCON BETWEEN " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ""
         End If
      End If
   Else
      If cmb_Situac.ItemData(cmb_Situac.ListIndex) <> 0 Then
         g_str_Parame = g_str_Parame & " WHERE LENGTH(TRIM(A.POSMAE_CODCON)) > 0 AND POSMAE_FECCON > 0 AND POSMAE_SITUAC =  '" & Format(cmb_Situac.ItemData(cmb_Situac.ListIndex), "000000") & "' "
         If chk_FecCon.Value = 0 Then
            g_str_Parame = g_str_Parame & " AND POSMAE_FECCON BETWEEN " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ""
         End If
      Else
         g_str_Parame = g_str_Parame & " WHERE LENGTH(TRIM(A.POSMAE_CODCON)) > 0 AND POSMAE_FECCON > 0 "
         If chk_FecCon.Value = 0 Then
            g_str_Parame = g_str_Parame & " AND POSMAE_FECCON BETWEEN " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ""
         End If
      End If
   End If
   If chk_ConHip.Value = 0 Then
      g_str_Parame = g_str_Parame & " AND POSMAE_CODCON = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "' "
   End If

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = g_rst_Princi!PROCLI_NUMDOC
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!PROCLI_APEPAT) & " " & Trim(g_rst_Princi!PROCLI_APEMAT) & " " & Trim(g_rst_Princi!PROCLI_NOMBRE)
      
      grd_Listad.Col = 2
      grd_Listad.Text = "" & g_rst_Princi!POSMAE_PROYEC
      
      grd_Listad.Col = 3
      If Len(Trim(g_rst_Princi!PROCLI_NUMTEL)) > 0 Then
         grd_Listad.Text = "" & g_rst_Princi!PROCLI_NUMTEL
      Else
         grd_Listad.Text = "" & g_rst_Princi!PROCLI_NUMCEL
      End If
      
      grd_Listad.Col = 4
      grd_Listad.Text = Right(CStr(g_rst_Princi!POSMAE_FECCON), 2) & "/" & Mid(CStr(g_rst_Princi!POSMAE_FECCON), 5, 2) & "/" & Left(CStr(g_rst_Princi!POSMAE_FECCON), 4)
      
      grd_Listad.Col = 5
      grd_Listad.Text = "" & Trim(g_rst_Princi!PARDES_DESCRI)
      
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!POSMAE_CODCON)
      
      grd_Listad.Col = 7
      grd_Listad.Text = g_rst_Princi!POSMAE_NUMSOL
      
      g_rst_Princi.MoveNext
   Loop
    
   grd_Listad.Redraw = True
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
    
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_ExpExc.Enabled = True
      grd_Listad.Enabled = True
   End If
    
   grd_Listad.ColWidth(7) = 0
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_int_FlgGrb = 2
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
    
   grd_Listad.Col = 7
   moddat_g_str_CodIte = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 2
   frm_IngCliPot_02.SSPanel7.Caption = frm_IngCliPot_02.SSPanel7.Caption & " - Modificar"
   frm_IngCliPot_02.Show 1
   
   Screen.MousePointer = 11
   Call cmd_Buscar_Click  'fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "300")
   cmb_Situac.AddItem "TODOS"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 0
   cmb_Situac.ListIndex = 0
   
   Call fs_Inicia
   chk_ConHip.Value = 1
   Call cmd_Buscar_Click
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1230
   grd_Listad.ColWidth(1) = 4920
   grd_Listad.ColWidth(2) = 3230
   grd_Listad.ColWidth(3) = 1110
   grd_Listad.ColWidth(4) = 1080
   grd_Listad.ColWidth(5) = 1320
   grd_Listad.ColWidth(6) = 1650
   grd_Listad.ColWidth(7) = 1
   grd_Listad.ColWidth(8) = 1
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   
   ipp_FecIni.Text = Format(date - 180)
   ipp_FecIni.Text = Format(Day(date), "00") & "/" & Format(Month(ipp_FecIni.Text), "00") & "/" & Year(ipp_FecIni.Text)
   ipp_FecFin.Text = (date)
End Sub

Public Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_ExpExc.Enabled = False
   grd_Listad.Enabled = False
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT POSMAE_NUMSOL, POSMAE_CODCON, PROCLI_TIPDOC, PROCLI_NUMDOC, PROCLI_APEPAT, PROCLI_APEMAT, "
   g_str_Parame = g_str_Parame & "       PROCLI_NOMBRE, POSMAE_FECCON, POSMAE_PROYEC, PROCLI_NUMTEL, PROCLI_NUMCEL, PARDES_DESCRI "
   g_str_Parame = g_str_Parame & "  FROM CRE_POSMAE A"
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PROCLI B ON B.PROCLI_TIPDOC = A.POSMAE_TIPDOC AND B.PROCLI_NUMDOC = A.POSMAE_NUMDOC "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = '300' AND D.PARDES_CODITE = A.POSMAE_SITUAC "
   If modgen_g_int_TipUsu = 20121 Then   'Si Tipo de Usuario es Consejero Hipotecario
      g_str_Parame = g_str_Parame & " WHERE A.POSMAE_CODCON = '" & modgen_g_str_CodUsu & "' AND A.POSMAE_FECCON > 0 AND A.POSMAE_SITUAC = '000001' "
   Else
      g_str_Parame = g_str_Parame & " WHERE LENGTH(TRIM(A.POSMAE_CODCON)) > 0 AND A.POSMAE_FECCON > 0 AND A.POSMAE_SITUAC = '000001' "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = g_rst_Princi!PROCLI_NUMDOC
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!PROCLI_APEPAT) & " " & Trim(g_rst_Princi!PROCLI_APEMAT) & " " & Trim(g_rst_Princi!PROCLI_NOMBRE)
      
      grd_Listad.Col = 2
      grd_Listad.Text = "" & g_rst_Princi!POSMAE_PROYEC
      
      grd_Listad.Col = 3
      If Len(Trim(g_rst_Princi!PROCLI_NUMTEL)) > 0 Then
         grd_Listad.Text = "" & g_rst_Princi!PROCLI_NUMTEL
      Else
         grd_Listad.Text = "" & g_rst_Princi!PROCLI_NUMCEL
      End If
      
      grd_Listad.Col = 4
      grd_Listad.Text = Right(CStr(g_rst_Princi!POSMAE_FECCON), 2) & "/" & Mid(CStr(g_rst_Princi!POSMAE_FECCON), 5, 2) & "/" & Left(CStr(g_rst_Princi!POSMAE_FECCON), 4)
      
      grd_Listad.Col = 5
      grd_Listad.Text = "" & Trim(g_rst_Princi!PARDES_DESCRI)
      
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!POSMAE_CODCON)
      
      grd_Listad.Col = 7
      grd_Listad.Text = g_rst_Princi!POSMAE_NUMSOL
      
      grd_Listad.Col = 8
      grd_Listad.Text = g_rst_Princi!POSMAE_FECCON
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
    
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_ExpExc.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   grd_Listad.ColWidth(7) = 0
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
   cmb_Situac.ListIndex = 0
End Sub

Private Sub grd_Listad_DblClick()
    Call cmd_Editar_Click
End Sub

Private Sub fs_GenExc()
Dim r_str_obs        As String
Dim r_int_ConAux     As Long
Dim r_int_Contad     As Long
Dim r_obj_Excel      As Excel.Application
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT POSMAE_NUMSOL, POSMAE_CODCON, PROCLI_TIPDOC, PROCLI_NUMDOC, PROCLI_APEPAT, PROCLI_APEMAT, PROCLI_NOMBRE,"
   g_str_Parame = g_str_Parame & "       POSMAE_FECCON, POSMAE_PROYEC, PROCLI_NUMTEL, PROCLI_NUMCEL, PARDES_DESCRI, POSMAE_COMMENT, POSMAE_PROMOT, POSMAE_CONSTR"
   g_str_Parame = g_str_Parame & "  FROM CRE_POSMAE A"
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PROCLI B ON B.PROCLI_TIPDOC = A.POSMAE_TIPDOC AND B.PROCLI_NUMDOC = A.POSMAE_NUMDOC "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = '300' AND D.PARDES_CODITE = A.POSMAE_SITUAC "
   If modgen_g_int_TipUsu = 20121 Then  'Si Tipo de Usuario es Consejero Hipotecario
      If cmb_Situac.ItemData(cmb_Situac.ListIndex) <> 0 Then
         g_str_Parame = g_str_Parame & " WHERE A.POSMAE_CODCON = '" & modgen_g_str_CodUsu & "' AND POSMAE_FECCON > 0 AND POSMAE_SITUAC = '" & Format(cmb_Situac.ItemData(cmb_Situac.ListIndex), "000000") & "' "
      Else
         g_str_Parame = g_str_Parame & " WHERE A.POSMAE_CODCON = '" & modgen_g_str_CodUsu & "' AND POSMAE_FECCON > 0 "
      End If
   Else
      If cmb_Situac.ItemData(cmb_Situac.ListIndex) <> 0 Then
         g_str_Parame = g_str_Parame & " WHERE LENGTH(TRIM(A.POSMAE_CODCON)) > 0 AND POSMAE_FECCON > 0 AND POSMAE_SITUAC = '" & Format(cmb_Situac.ItemData(cmb_Situac.ListIndex), "000000") & "' "
      Else
         g_str_Parame = g_str_Parame & " WHERE LENGTH(TRIM(A.POSMAE_CODCON)) > 0 AND POSMAE_FECCON > 0 "
      End If
   End If
   
   If chk_FecCon.Value = 0 Then
      g_str_Parame = g_str_Parame & " AND A.POSMAE_FECCON BETWEEN " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ""
   End If

   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
    
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      'Unir celdas
      .Range("B5") = "LISTADO DE CLIENTES POTENCIALES"
      .Range("B5:D5").Font.Underline = True
      .Range("B5:D5").Font.Bold = True
      .Range("A5:K5").Merge
      
      .Cells(2, 11) = "Dpto. Tecnología e Informática"
      .Cells(3, 11) = "Desarrollo de Sistemas"
      .Cells(7, 1) = "ITEM"
      .Cells(7, 2) = "N.DOCUMENTO"
      .Cells(7, 3) = "APELLIDOS Y NOMBRES"
      .Cells(7, 4) = "PROYECTO"
      .Cells(7, 5) = "TELEFONO"
      .Cells(7, 6) = "F. REGISTRO"
      .Cells(7, 7) = "SITUACION"
      .Cells(7, 8) = "CONSEJERO"
      .Cells(7, 9) = "COMENTARIOS"
      .Cells(7, 10) = "PROMOTOR"
      .Cells(7, 11) = "CONSTRUCTOR"
      
      .Columns("A").ColumnWidth = 6
      .Columns("B").ColumnWidth = 12
      .Columns("C").ColumnWidth = 35
      .Columns("D").ColumnWidth = 30
      .Columns("E").ColumnWidth = 10
      .Columns("F").ColumnWidth = 10
      .Columns("G").ColumnWidth = 10
      .Columns("H").ColumnWidth = 14
      .Columns("I").ColumnWidth = 50
      .Columns("J").ColumnWidth = 40
      .Columns("K").ColumnWidth = 40
      .Cells(7, 1).Font.Bold = True
      .Cells(7, 2).Font.Bold = True
      .Cells(7, 3).Font.Bold = True
      .Cells(7, 4).Font.Bold = True
      .Cells(7, 5).Font.Bold = True
      .Cells(7, 6).Font.Bold = True
      .Cells(7, 7).Font.Bold = True
      .Cells(7, 8).Font.Bold = True
      .Cells(7, 9).Font.Bold = True
      .Cells(7, 10).Font.Bold = True
      .Cells(7, 11).Font.Bold = True
      .Range("B7").Font.Bold = True
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Range("A5").HorizontalAlignment = xlHAlignCenter
      .Cells(2, 9).HorizontalAlignment = xlHAlignRight
      .Cells(3, 9).HorizontalAlignment = xlHAlignRight
      
      .Range("A1:K1000").Font.Name = "Arial"
      .Range("A1:K1000").Font.Size = 8
   End With
    
   r_int_ConAux = 8
   r_int_Contad = 1
   g_rst_Princi.MoveFirst
    
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad) = Format(r_int_ConAux - 7, "0000")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad + 1) = g_rst_Princi!PROCLI_NUMDOC
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad + 2) = Trim(g_rst_Princi!PROCLI_APEPAT) & " " & Trim(g_rst_Princi!PROCLI_APEMAT) & " " & Trim(g_rst_Princi!PROCLI_NOMBRE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad + 3) = Trim(g_rst_Princi!POSMAE_PROYEC)
      If Len(Trim(g_rst_Princi!PROCLI_NUMTEL)) > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad + 4) = g_rst_Princi!PROCLI_NUMTEL
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad + 4) = g_rst_Princi!PROCLI_NUMCEL
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad + 5) = "'" & Right(CStr(g_rst_Princi!POSMAE_FECCON), 2) & "/" & Mid(CStr(g_rst_Princi!POSMAE_FECCON), 5, 2) & "/" & Left(CStr(g_rst_Princi!POSMAE_FECCON), 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad + 6) = Trim(g_rst_Princi!PARDES_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad + 7) = Trim(g_rst_Princi!POSMAE_CODCON)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad + 8) = Trim(g_rst_Princi!POSMAE_COMMENT)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad + 9) = Trim(g_rst_Princi!POSMAE_PROMOT)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, r_int_Contad + 10) = Trim(g_rst_Princi!POSMAE_CONSTR)
      
      r_int_ConAux = r_int_ConAux + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Ape_Nom_Click()
   If Len(Trim(pnl_Ape_Nom.Tag)) = 0 Or pnl_Ape_Nom.Tag = "D" Then
      pnl_Ape_Nom.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Ape_Nom.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Consej_Click()
   If Len(Trim(pnl_Consej.Tag)) = 0 Or pnl_Consej.Tag = "D" Then
      pnl_Consej.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Consej.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

Private Sub pnl_FecReg_Click()
   If Len(Trim(pnl_FecReg.Tag)) = 0 Or pnl_FecReg.Tag = "D" Then
      pnl_FecReg.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "C")
   Else
      pnl_FecReg.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "C-")
   End If
End Sub

Private Sub pnl_NumDoc_Click()
   If Len(Trim(pnl_NumDoc.Tag)) = 0 Or pnl_NumDoc.Tag = "D" Then
      pnl_NumDoc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_NumDoc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Proyec_Click()
   If Len(Trim(pnl_Proyec.Tag)) = 0 Or pnl_Proyec.Tag = "D" Then
      pnl_Proyec.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Proyec.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Situac_Click()
   If Len(Trim(pnl_Situac.Tag)) = 0 Or pnl_Situac.Tag = "D" Then
      pnl_Situac.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Situac.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Telefo_Click()
   If Len(Trim(pnl_Telefo.Tag)) = 0 Or pnl_Telefo.Tag = "D" Then
      pnl_Telefo.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Telefo.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub
