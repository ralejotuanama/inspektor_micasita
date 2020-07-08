VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_IngCliPot_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   Icon            =   "AteCliPot_frm_002.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel4 
      Height          =   675
      Left            =   30
      TabIndex        =   0
      Top             =   765
      Width           =   8025
      _Version        =   65536
      _ExtentX        =   14155
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
      Begin VB.CommandButton cmd_GrabarProsp 
         Height          =   600
         Left            =   1800
         Picture         =   "AteCliPot_frm_002.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Grabar Cliente Prospecto"
         Top             =   45
         Width           =   600
      End
      Begin VB.CommandButton cmd_Grabar 
         Height          =   600
         Left            =   1215
         Picture         =   "AteCliPot_frm_002.frx":08D6
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Grabar Cliente Potencial"
         Top             =   45
         Width           =   600
      End
      Begin VB.CommandButton cmd_Limpia 
         Height          =   600
         Left            =   630
         Picture         =   "AteCliPot_frm_002.frx":0D18
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Limpiar Datos"
         Top             =   45
         Width           =   600
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   600
         Left            =   7380
         Picture         =   "AteCliPot_frm_002.frx":1022
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Salir de la Opción"
         Top             =   45
         Width           =   600
      End
      Begin VB.CommandButton cmd_Buscar 
         Height          =   600
         Left            =   45
         Picture         =   "AteCliPot_frm_002.frx":1464
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar Datos"
         Top             =   45
         Width           =   600
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   4920
      Left            =   30
      TabIndex        =   23
      Top             =   2280
      Width           =   8025
      _Version        =   65536
      _ExtentX        =   14155
      _ExtentY        =   8678
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
      Begin VB.ComboBox cmb_ConHip 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   4500
         Width           =   5880
      End
      Begin VB.TextBox txt_Nombre 
         Height          =   330
         Left            =   1890
         MaxLength       =   30
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   810
         Width           =   3345
      End
      Begin VB.TextBox txt_Constru 
         Height          =   330
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2610
         Width           =   5865
      End
      Begin VB.TextBox txt_Promo 
         Height          =   330
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2250
         Width           =   5865
      End
      Begin VB.TextBox txt_Comment 
         Height          =   735
         Left            =   1890
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   3735
         Width           =   5865
      End
      Begin VB.ComboBox cmb_Situac 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3375
         Width           =   3390
      End
      Begin VB.TextBox txt_numCel 
         Height          =   330
         Left            =   6210
         MaxLength       =   9
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1125
         Width           =   1545
      End
      Begin VB.TextBox txt_Proyec 
         Height          =   330
         Left            =   1890
         MaxLength       =   100
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1890
         Width           =   5865
      End
      Begin VB.TextBox txt_DirEle 
         Height          =   330
         Left            =   1890
         MaxLength       =   120
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1530
         Width           =   3345
      End
      Begin VB.TextBox txt_numTel 
         Height          =   330
         Left            =   1890
         MaxLength       =   12
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1170
         Width           =   1500
      End
      Begin VB.TextBox txt_ApePat 
         Height          =   330
         Left            =   1890
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   90
         Width           =   3345
      End
      Begin VB.TextBox txt_ApeMat 
         Height          =   330
         Left            =   1890
         MaxLength       =   30
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   450
         Width           =   3345
      End
      Begin EditLib.fpDoubleSingle ipp_ValInm 
         Height          =   330
         Left            =   1890
         TabIndex        =   13
         Top             =   3015
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
      Begin EditLib.fpDateTime ipp_FecCon 
         Height          =   330
         Left            =   6255
         TabIndex        =   14
         Top             =   3015
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
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
      Begin VB.Label Label13 
         Caption         =   "Consejero :"
         Height          =   330
         Left            =   135
         TabIndex        =   40
         Top             =   4500
         Width           =   1635
      End
      Begin VB.Label Label12 
         Caption         =   "Constructor:"
         Height          =   285
         Left            =   135
         TabIndex        =   34
         Top             =   2610
         Width           =   1635
      End
      Begin VB.Label Label11 
         Caption         =   "Promotor:"
         Height          =   285
         Left            =   135
         TabIndex        =   33
         Top             =   2250
         Width           =   1635
      End
      Begin VB.Label Label10 
         Caption         =   "Comentarios:"
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   3735
         Width           =   1605
      End
      Begin VB.Label Label9 
         Caption         =   "Situacion:"
         Height          =   375
         Left            =   135
         TabIndex        =   37
         Top             =   3375
         Width           =   1635
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha de Registro:"
         Height          =   270
         Left            =   4695
         TabIndex        =   36
         Top             =   3015
         Width           =   1665
      End
      Begin VB.Label Label7 
         Caption         =   "Celular:"
         Height          =   285
         Left            =   5490
         TabIndex        =   30
         Top             =   1170
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "Proyecto:"
         Height          =   285
         Left            =   135
         TabIndex        =   32
         Top             =   1890
         Width           =   1635
      End
      Begin VB.Label lbl_General 
         Caption         =   "Valor Inmueble:"
         Height          =   285
         Index           =   61
         Left            =   135
         TabIndex        =   35
         Top             =   3015
         Width           =   1635
      End
      Begin VB.Label Label17 
         Caption         =   "E-mail:"
         Height          =   285
         Left            =   135
         TabIndex        =   31
         Top             =   1530
         Width           =   1635
      End
      Begin VB.Label Label16 
         Caption         =   "Teléfono :"
         Height          =   285
         Left            =   135
         TabIndex        =   29
         Top             =   1170
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "Apellido Paterno:"
         Height          =   285
         Left            =   135
         TabIndex        =   26
         Top             =   90
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "Apellido Materno:"
         Height          =   285
         Left            =   135
         TabIndex        =   27
         Top             =   450
         Width           =   1635
      End
      Begin VB.Label Label5 
         Caption         =   "Nombres:"
         Height          =   285
         Left            =   135
         TabIndex        =   28
         Top             =   810
         Width           =   1635
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   780
      Left            =   30
      TabIndex        =   39
      Top             =   1470
      Width           =   8025
      _Version        =   65536
      _ExtentX        =   14155
      _ExtentY        =   1376
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
         Left            =   1890
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   400
         Width           =   2775
      End
      Begin VB.ComboBox cmb_TipDoc 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   60
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Nro. Doc. Id.:"
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Docum. Identidad:"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   75
         Width           =   1785
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   720
      Left            =   30
      TabIndex        =   21
      Top             =   30
      Width           =   8025
      _Version        =   65536
      _ExtentX        =   14155
      _ExtentY        =   1270
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
         TabIndex        =   22
         Top             =   60
         Width           =   5445
         _Version        =   65536
         _ExtentX        =   9604
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "Mantenimiento de Cliente Potencial"
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
      End
      Begin MSMAPI.MAPISession mps_Sesion 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   60
         Picture         =   "AteCliPot_frm_002.frx":176E
         Top             =   60
         Width           =   480
      End
   End
End
Attribute VB_Name = "frm_IngCliPot_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_ConHip()      As moddat_tpo_Genera







Private Sub cmb_ConHip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(cmd_Grabar)
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
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      Case 1
         If Len(Trim(txt_NumDoc.Text)) <> 8 Then
            MsgBox "El DNI debe tener 8 caracteres, favor verificar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumDoc)
            Exit Sub
         End If
      Case Else
         If Len(Trim(txt_NumDoc.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumDoc)
            Exit Sub
         End If
   End Select
   
   'Verificando que Cliente no haya sido ingresado como Cliente Negativo o PEP
   If Not atecli_gf_Buscar_BasNeg(CInt(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)), Trim(txt_NumDoc.Text)) Then
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT PROCLI_TIPDOC,PROCLI_NUMDOC,PROCLI_APEPAT,PROCLI_APEMAT, "
   g_str_Parame = g_str_Parame & "       PROCLI_NOMBRE,PROCLI_NUMTEL,PROCLI_NUMCEL,PROCLI_CORREO  "
   g_str_Parame = g_str_Parame & "  FROM CRE_PROCLI "
   g_str_Parame = g_str_Parame & " WHERE PROCLI_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " "
   g_str_Parame = g_str_Parame & "   AND PROCLI_NUMDOC = " & Trim(txt_NumDoc.Text) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Call fs_Activa(True)
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      cmd_Buscar.Enabled = False
      cmb_ConHip.Enabled = False
      Call gs_BuscarCombo_Item(cmb_Situac, 1)
      Call gs_SetFocus(txt_ApePat)
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT POSMAE_CODCON "
      g_str_Parame = g_str_Parame & "  FROM CRE_POSMAE "
      g_str_Parame = g_str_Parame & " WHERE POSMAE_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " "
      g_str_Parame = g_str_Parame & "   AND POSMAE_NUMDOC = " & Trim(txt_NumDoc.Text) & " "
      g_str_Parame = g_str_Parame & "   AND POSMAE_SITUAC = '000001' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         MsgBox "El Cliente ya está registrado por el consejero: " & Trim(g_rst_Genera!POSMAE_CODCON), vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      Else
         moddat_g_int_FlgGrb = 3
         txt_ApePat.Text = IIf(IsNull(g_rst_Princi!PROCLI_APEPAT), "", g_rst_Princi!PROCLI_APEPAT)
         txt_ApeMat.Text = IIf(IsNull(g_rst_Princi!PROCLI_APEMAT), "", g_rst_Princi!PROCLI_APEMAT)
         txt_Nombre.Text = IIf(IsNull(g_rst_Princi!PROCLI_NOMBRE), "", g_rst_Princi!PROCLI_NOMBRE)
         txt_numTel.Text = IIf(IsNull(g_rst_Princi!PROCLI_NUMTEL), "", g_rst_Princi!PROCLI_NUMTEL)
         txt_numCel.Text = IIf(IsNull(g_rst_Princi!PROCLI_NUMCEL), "", g_rst_Princi!PROCLI_NUMCEL)
         txt_DirEle.Text = IIf(IsNull(g_rst_Princi!PROCLI_CORREO), "", g_rst_Princi!PROCLI_CORREO)
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      Call fs_Activa(True)
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      cmd_Buscar.Enabled = False
      Call gs_SetFocus(txt_ApePat)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_Grabar_Click()

      ' ********************************* 15012020 INICIO BY RAT
             Dim a As Integer
              
       a = validacionclientexproblemas(txt_NumDoc.Text, "")
         
           Dim var As Integer
           var = 110
       
      
           g_str_Parame = ""
             g_str_Parame = "USP_CRE_INSPEK ("
              g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
             g_str_Parame = g_str_Parame & "'" & var & "', "
             g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
             g_str_Parame = g_str_Parame & "'" & Trim(modgen_g_str_rptwebservice) & "', "
             g_str_Parame = g_str_Parame & "'" & Trim(txt_Nombre) & "', "
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
             g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                   
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            Debug.Print ("invalido");
            Else
            Debug.Print ("valido");
            modgen_g_str_nombreformulario = ""
            End If
           
            If a <> 0 Then
              Debug.Print "CLIENTE NO TIENE PROBLEMAS"
            Else
              Debug.Print "CLIENTE  TIENE PROBLEMAS"
              MsgBox (modgen_g_str_rptwebservice)
              Call fs_Limpia
              Call fs_Activa(False)
              cmb_TipDoc.Enabled = True
              txt_NumDoc.Enabled = True
              cmd_Buscar.Enabled = True
              cmd_GrabarProsp.Enabled = False
              Exit Sub
            End If
       ' ********************************* 15012020 FIN BY RAT
       
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
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
   'If Len(Trim(txt_numTel.Text)) = 0 And Len(Trim(txt_numCel.Text)) = 0 Then
   '   MsgBox "Debe ingresar el Teléfono o Celular.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(txt_numTel)
   '   Exit Sub
   'End If
   
   'MsgBox (Trim(txt_Nombre.Text) & " " & Trim(txt_ApePat.Text) & " " & Trim(txt_ApeMat.Text))
   
   If Len(Trim(txt_numTel.Text)) > 0 Then
      If Len(Trim(txt_numTel.Text)) < 6 Then
         MsgBox "El Teléfono debe tener al menos 6 digitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_numTel)
         Exit Sub
      End If
   End If
   If Len(Trim(txt_numCel.Text)) > 0 Then
      If Len(Trim(txt_numCel.Text)) < 8 Then
         MsgBox "El Celular debe tener al menos 8 digitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_numCel)
         Exit Sub
      End If
   End If
   If Len(Trim(txt_DirEle.Text)) > 0 Then
      If Not gf_ValidarEmail(txt_DirEle.Text) Then
         MsgBox "El E-mail no tiene el formato correcto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle)
         Exit Sub
      End If
   End If
   'If Len(Trim(txt_Proyec.Text)) = 0 Then
   '   MsgBox "Debe ingresar el Proyecto.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(txt_Proyec)
   '   Exit Sub
   'End If
   If CDate(ipp_FecCon.Text) > Format(date, "dd/mm/yyyy") Then
      MsgBox "Fecha no puede ser mayor a la del día actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecCon)
      Exit Sub
   End If
   'If Len(Trim(ipp_ValInm.Text)) = 0 Or Val(ipp_ValInm) = 0 Or ipp_ValInm.Value = 0 Then
   '   MsgBox "Debe ingresar el Valor del Inmueble.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(ipp_ValInm)
   '   Exit Sub
   'End If
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situacion.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   If cmb_ConHip.ListIndex = -1 Then
      MsgBox "Debe seleccionar Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ConHip)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT EJECMC_CODEJE  "
   g_str_Parame = g_str_Parame & " FROM CRE_EJECMC"
   g_str_Parame = g_str_Parame & " WHERE RTRIM(EJECMC_APEPAT) || ' ' || RTRIM(EJECMC_APEMAT) || ' ' || RTRIM(EJECMC_NOMBRE) = "
   g_str_Parame = g_str_Parame & " '" & cmb_ConHip.Text & "'"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   If moddat_g_int_FlgGrb = 1 Then
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = ""
         g_str_Parame = "USP_CRE_POSMAE_POT ("
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EJECMC_CODEJE) & "', "
         g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Proyec.Text & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Promo.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Constru.Text) & "', "
         g_str_Parame = g_str_Parame & Format(ipp_ValInm.Text, "############0.00") & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecCon.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & Format(cmb_Situac.ItemData(cmb_Situac.ListIndex), "000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Comment.Text) & "' , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & 1 & ")"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
            moddat_g_int_CntErr = 0
         End If
         
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento (USP_CRE_POSMAE_POT). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_FlgGOK = False
            End If
         End If
      Loop
            
      moddat_g_int_FlgGOK = False
      If moddat_g_int_CntErr = 0 Then
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = ""
            g_str_Parame = "USP_CRE_PROCLI ("
            g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
            g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
            g_str_Parame = g_str_Parame & "'" & IIf(txt_numTel.Text = "", "", txt_numTel.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & IIf(txt_numCel.Text = "", "", txt_numCel.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & IIf(txt_DirEle.Text = "", "", txt_DirEle.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Ejecutable
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Terminal
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & 1 & ")"
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
            
            If moddat_g_int_CntErr > 0 Then
               If MsgBox("No se pudo completar el procedimiento (USP_CRE_PROCLI). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_FlgGOK = True
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      End If
   
   ElseIf moddat_g_int_FlgGrb = 2 Then
   
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = ""
         g_str_Parame = "USP_CRE_POSMAE_POT ("
         g_str_Parame = g_str_Parame & moddat_g_str_CodIte & ", "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EJECMC_CODEJE) & "', "
         g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Proyec.Text & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Promo.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Constru.Text) & "', "
         g_str_Parame = g_str_Parame & Format(ipp_ValInm.Text, "############0.00") & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecCon.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & Format(cmb_Situac.ItemData(cmb_Situac.ListIndex), "000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Comment.Text) & "' , "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
         g_str_Parame = g_str_Parame & 2 & ")"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento (USP_CRE_POSMAE_POT). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_FlgGOK = True
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   
      If moddat_g_int_CntErr = 0 Then
         moddat_g_int_FlgGOK = False
    
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = ""
            g_str_Parame = "USP_CRE_PROCLI ("
            g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
            g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
            g_str_Parame = g_str_Parame & "'" & IIf(txt_numTel.Text = "", "", txt_numTel.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & IIf(txt_numCel.Text = "", "", txt_numCel.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & IIf(txt_DirEle.Text = "", "", txt_DirEle.Text) & "', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Ejecutable
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Terminal
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & 2 & ")"
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
            If moddat_g_int_CntErr > 0 Then
               If MsgBox("No se pudo completar el procedimiento (USP_CRE_PROCLI). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_FlgGOK = True
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      End If
   
   Else
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = ""
         g_str_Parame = "USP_CRE_POSMAE_POT ("
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EJECMC_CODEJE) & "', "
         g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Proyec.Text & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Promo.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Constru.Text) & "', "
         g_str_Parame = g_str_Parame & Format(ipp_ValInm.Text, "############0.00") & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecCon.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & Format(cmb_Situac.ItemData(cmb_Situac.ListIndex), "000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Comment.Text) & "' , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & 1 & ")"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
            moddat_g_int_CntErr = 0
         End If
         
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento (USP_CRE_POSMAE_POT). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_FlgGOK = False
            End If
         End If
      Loop

   
      If moddat_g_int_CntErr = 0 Then
         moddat_g_int_FlgGOK = False
    
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = ""
            g_str_Parame = "USP_CRE_PROCLI ("
            g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
            g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
            g_str_Parame = g_str_Parame & "'" & IIf(txt_numTel.Text = "", "", txt_numTel.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & IIf(txt_numCel.Text = "", "", txt_numCel.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & IIf(txt_DirEle.Text = "", "", txt_DirEle.Text) & "', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Ejecutable
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Terminal
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & 2 & ")"
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
            If moddat_g_int_CntErr > 0 Then
               If MsgBox("No se pudo completar el procedimiento (USP_CRE_PROCLI). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_FlgGOK = True
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      End If

   End If
   
   cmd_Grabar.Enabled = False
   cmd_GrabarProsp.Enabled = True
   
'   Call fs_Activa(False)
'   Call fs_Limpia
'   Unload Me
End Sub

Private Sub cmd_GrabarProsp_Click()
   
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
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
   If Len(Trim(txt_numTel.Text)) = 0 And Len(Trim(txt_numCel.Text)) = 0 Then
      MsgBox "Debe ingresar el Teléfono o Celular.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_numTel)
      Exit Sub
   End If
   If Len(Trim(txt_numTel.Text)) > 0 Then
      If Len(Trim(txt_numTel.Text)) < 6 Then
         MsgBox "El Teléfono debe tener al menos 6 digitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_numTel)
         Exit Sub
      End If
   End If
   If Len(Trim(txt_numCel.Text)) > 0 Then
      If Len(Trim(txt_numCel.Text)) < 8 Then
         MsgBox "El Celular debe tener al menos 8 digitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_numCel)
         Exit Sub
      End If
   End If
   If Len(Trim(txt_DirEle.Text)) < 1 Then
      MsgBox "El E-mail no tiene el formato correcto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DirEle)
      Exit Sub
   Else
      If Not gf_ValidarEmail(txt_DirEle.Text) Then
         MsgBox "El E-mail no tiene el formato correcto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle)
         Exit Sub
      End If
   End If
   If CDate(ipp_FecCon.Text) > Format(date, "dd/mm/yyyy") Then
      MsgBox "Fecha no puede ser mayor a la del día actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecCon)
      Exit Sub
   End If
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situacion.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   If cmb_ConHip.ListIndex = -1 Then
      MsgBox "Debe seleccionar Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ConHip)
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PROMAE_NUMSOL  "
   g_str_Parame = g_str_Parame & " FROM CRE_PROMAE"
   g_str_Parame = g_str_Parame & " WHERE PROMAE_TIPDOC='" & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & "'"
   g_str_Parame = g_str_Parame & " AND PROMAE_NUMDOC='" & Trim(txt_NumDoc.Text) & "'"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      MsgBox "El Cliente ya se encuentra registrado tambien en el " & Chr(13) & _
             "formulario de 'Mantenimiento de Prospecto'.   ", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT EJECMC_CODEJE  "
   g_str_Parame = g_str_Parame & " FROM CRE_EJECMC"
   g_str_Parame = g_str_Parame & " WHERE RTRIM(EJECMC_APEPAT) || ' ' || RTRIM(EJECMC_APEMAT) || ' ' || RTRIM(EJECMC_NOMBRE) = "
   g_str_Parame = g_str_Parame & " '" & cmb_ConHip.Text & "'"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   If moddat_g_int_FlgGrb = 1 Then
      Do While moddat_g_int_FlgGOK = False
               
         g_str_Parame = ""
         g_str_Parame = "USP_CRE_PROMAE ("
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EJECMC_CODEJE) & "', "
         g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Proyec.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Promo.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Constru.Text & "', "
         g_str_Parame = g_str_Parame & Format(ipp_ValInm.Text, "############0.00") & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecCon.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & Format(cmb_Situac.ItemData(cmb_Situac.ListIndex), "000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Comment.Text) & "' , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & 1 & ")"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
            moddat_g_int_CntErr = 0
         End If
         
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento (USP_CRE_PROMAE). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_FlgGOK = False
            End If
         End If
      Loop
            
      moddat_g_int_FlgGOK = False
      If moddat_g_int_CntErr = 0 Then
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = ""
            g_str_Parame = "USP_CRE_PROCLI ("
            g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
            g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_numTel.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_numCel.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_DirEle.Text & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Terminal
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Ejecutable
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & 2 & ")"
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
            
            If moddat_g_int_CntErr > 0 Then
               If MsgBox("No se pudo completar el procedimiento (USP_CRE_PROCLI). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_FlgGOK = True
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      End If
   ElseIf moddat_g_int_FlgGrb = 2 Then
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = ""
         g_str_Parame = "USP_CRE_PROMAE ("
         g_str_Parame = g_str_Parame & moddat_g_str_CodIte & ", "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EJECMC_CODEJE) & "', "
         g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Proyec.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Promo.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Constru.Text & "', "
         g_str_Parame = g_str_Parame & Format(ipp_ValInm.Text, "############0.00") & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecCon.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & Format(cmb_Situac.ItemData(cmb_Situac.ListIndex), "000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Comment.Text) & "' , "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
         g_str_Parame = g_str_Parame & 1 & ")"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento (USP_CRE_PROMAE). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_FlgGOK = True
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   
      If moddat_g_int_CntErr = 0 Then
         moddat_g_int_FlgGOK = False
    
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = ""
            g_str_Parame = "USP_CRE_PROCLI ("
            g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
            g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_numTel.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_numCel.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_DirEle.Text & "', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Terminal
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Ejecutable
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & 2 & ")"
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
            If moddat_g_int_CntErr > 0 Then
               If MsgBox("No se pudo completar el procedimiento (USP_CRE_PROCLI). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_FlgGOK = True
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      End If
      
   Else
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = ""
         g_str_Parame = "USP_CRE_PROMAE ("
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EJECMC_CODEJE) & "', "
         g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Proyec.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Promo.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Constru.Text & "', "
         g_str_Parame = g_str_Parame & Format(ipp_ValInm.Text, "############0.00") & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecCon.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & Format(cmb_Situac.ItemData(cmb_Situac.ListIndex), "000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Comment.Text) & "' , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & 1 & ")"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
            moddat_g_int_CntErr = 0
         End If
         
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento (USP_CRE_PROMAE). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_FlgGOK = False
            End If
         End If
      Loop

      
      If moddat_g_int_CntErr = 0 Then
         moddat_g_int_FlgGOK = False
    
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = ""
            g_str_Parame = "USP_CRE_PROCLI ("
            g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " , "
            g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_numTel.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_numCel.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_DirEle.Text & "', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "       'Código Usuario
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "        'Nombre Terminal
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "       'Nombre Ejecutable
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & 2 & ")"
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
            If moddat_g_int_CntErr > 0 Then
               If MsgBox("No se pudo completar el procedimiento (USP_CRE_PROCLI). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_FlgGOK = True
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      End If
   End If
   
   Call fs_Activa(False)
   Call fs_Limpia
   Unload Me
End Sub

Private Sub cmd_Limpia_Click()
   If moddat_g_int_FlgGrb = 1 Then
      Call fs_Limpia
      Call fs_Activa(False)
      cmb_TipDoc.Enabled = True
      txt_NumDoc.Enabled = True
      cmd_Buscar.Enabled = True
      cmd_GrabarProsp.Enabled = False
   Else
      txt_ApePat.Text = ""
      txt_ApeMat.Text = ""
      txt_Nombre.Text = ""
      txt_numTel.Text = ""
      txt_numCel.Text = ""
      txt_DirEle.Text = ""
      txt_Proyec.Text = ""
      txt_Promo.Text = ""
      txt_Constru.Text = ""
      ipp_ValInm.Value = 0
      ipp_FecCon.Text = Format(date, "dd/mm/yyyy")
      txt_Comment.Text = ""
   End If
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Call gs_CentraForm(Me)
   Me.Caption = modgen_g_str_NomPlt
    
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "300")
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
   Call fs_Limpia
   Call fs_Activa(False)
   
   If moddat_g_int_FlgGrb = 2 Then
      Call fs_Buscar
   End If
   
   If modgen_g_int_TipUsu = 20121 Or modgen_g_int_TipUsu = 20111 Then
      cmb_ConHip.Enabled = False
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT RTRIM(EJECMC_APEPAT) || ' ' || RTRIM(EJECMC_APEMAT) || ' ' || RTRIM(EJECMC_NOMBRE) AS CONS "
   g_str_Parame = g_str_Parame & " FROM CRE_EJECMC A, CRE_EJETIP B"
   g_str_Parame = g_str_Parame & " WHERE EJETIP_CODEJE = EJECMC_CODEJE AND RTRIM(EJECMC_CODEJE)='" & Trim(modgen_g_str_CodUsu) & "' "
   g_str_Parame = g_str_Parame & " AND A.EJECMC_SITUAC = 1 AND B.EJETIP_TIPEJE=121"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.EOF And g_rst_Princi.BOF) Then
      cmb_ConHip.Text = g_rst_Princi!CONS
   End If
      
   cmd_GrabarProsp.Enabled = False
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   txt_numTel.Text = ""
   txt_numCel.Text = ""
   txt_DirEle.Text = ""
   txt_Proyec.Text = ""
   txt_Promo.Text = ""
   txt_Constru.Text = ""
   ipp_ValInm.Value = 0
   ipp_FecCon.Text = Format(date, "dd/mm/yyyy")
   cmb_Situac.ListIndex = -1
   txt_Comment.Text = ""
''   cmb_ConHip.ListIndex = -1
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmd_Grabar.Enabled = p_Activa
   txt_ApePat.Enabled = p_Activa
   txt_ApeMat.Enabled = p_Activa
   txt_Nombre.Enabled = p_Activa
   txt_numTel.Enabled = p_Activa
   txt_numCel.Enabled = p_Activa
   txt_DirEle.Enabled = p_Activa
   txt_Proyec.Enabled = p_Activa
   txt_Constru.Enabled = p_Activa
   txt_Promo.Enabled = p_Activa
   ipp_ValInm.Enabled = p_Activa
   ipp_FecCon.Enabled = p_Activa
   cmb_Situac.Enabled = p_Activa
   txt_Comment.Enabled = p_Activa
   cmb_ConHip.Enabled = p_Activa
End Sub

Public Sub fs_Buscar()
   Screen.MousePointer = 11
   cmd_Buscar.Enabled = False
   cmd_Limpia.Enabled = False
   Call fs_Limpia
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT POSMAE_NUMSOL, POSMAE_CODCON, PROCLI_TIPDOC, PROCLI_NUMDOC, PROCLI_APEPAT, PROCLI_APEMAT,"
   g_str_Parame = g_str_Parame & "       PROCLI_NOMBRE, PROCLI_NUMTEL, PROCLI_NUMCEL, PROCLI_CORREO, POSMAE_FECCON, "
   g_str_Parame = g_str_Parame & "       POSMAE_PROYEC, POSMAE_PROMOT, POSMAE_CONSTR, POSMAE_VALINM, POSMAE_SITUAC, POSMAE_COMMENT "
   g_str_Parame = g_str_Parame & "  FROM CRE_POSMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PROCLI B ON PROCLI_TIPDOC = POSMAE_TIPDOC AND PROCLI_NUMDOC = POSMAE_NUMDOC "
   g_str_Parame = g_str_Parame & " WHERE POSMAE_NUMSOL = " & moddat_g_str_CodIte & "   "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!PROCLI_TIPDOC)
   txt_NumDoc.Text = Trim(g_rst_Princi!PROCLI_NUMDOC)
   txt_ApePat.Text = Trim(g_rst_Princi!PROCLI_APEPAT)
   txt_ApeMat.Text = Trim(g_rst_Princi!PROCLI_APEMAT)
   txt_Nombre.Text = Trim(g_rst_Princi!PROCLI_NOMBRE)
   txt_numTel.Text = "" & Trim(g_rst_Princi!PROCLI_NUMTEL)
   txt_numCel.Text = "" & Trim(g_rst_Princi!PROCLI_NUMCEL)
   txt_DirEle.Text = "" & Trim(g_rst_Princi!PROCLI_CORREO)
   txt_Proyec.Text = "" & Trim(g_rst_Princi!POSMAE_PROYEC)
   txt_Promo.Text = "" & Trim(g_rst_Princi!POSMAE_PROMOT)
   txt_Constru.Text = "" & Trim(g_rst_Princi!POSMAE_CONSTR)
   ipp_ValInm.Text = Trim(g_rst_Princi!POSMAE_VALINM)
   ipp_FecCon.Text = Right(CStr(g_rst_Princi!POSMAE_FECCON), 2) & "/" & Mid(CStr(g_rst_Princi!POSMAE_FECCON), 5, 2) & "/" & Left(CStr(g_rst_Princi!POSMAE_FECCON), 4)
   Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Princi!POSMAE_SITUAC)
   txt_Comment.Text = "" & Trim(g_rst_Princi!POSMAE_COMMENT)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT RTRIM(EJECMC_APEPAT) || ' ' || RTRIM(EJECMC_APEMAT) || ' ' || RTRIM(EJECMC_NOMBRE) AS PROMAE_CONS "
   g_str_Parame = g_str_Parame & " FROM CRE_EJECMC"
   g_str_Parame = g_str_Parame & " WHERE RTRIM(EJECMC_CODEJE)='" & Trim(g_rst_Princi!POSMAE_CODCON) & "'"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.EOF And g_rst_Princi.BOF) Then
      cmb_ConHip.Text = Trim(g_rst_Princi!PROMAE_CONS)
   End If

   Call fs_Activa(True)
   txt_NumDoc.Enabled = False
   cmb_TipDoc.Enabled = False
   If modgen_g_int_TipUsu = 20121 Then          'Si Tipo de Usuario es Consejero Hipotecario
      ipp_FecCon.Enabled = False
   End If
   Call gs_SetFocus(txt_ApePat)
   Screen.MousePointer = 0
End Sub

Private Sub ipp_FecCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(cmb_Situac)
   End If
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(txt_Comment)
   End If
End Sub

Private Sub ipp_ValInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_FecCon.Enabled = True Then
         Call gs_SetFocus(ipp_FecCon)
      Else
         Call gs_SetFocus(cmb_Situac)
      End If
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(txt_Nombre)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " ")
   End If
End Sub

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(txt_ApeMat)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " ")
   End If
End Sub

Private Sub txt_Comment_GotFocus()
   txt_Comment.SelStart = Len(Right(txt_Comment.Text, Len(txt_Comment.Text)))
End Sub

Private Sub txt_Comment_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ConHip)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_CadOri & modgen_g_con_CadEnc & " ")
   End If
End Sub

Private Sub txt_Constru_GotFocus()
   Call gs_SelecTodo(txt_Constru)
End Sub

Private Sub txt_Constru_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValInm)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_CadOri & "- ")
   End If
End Sub

Private Sub txt_DirEle_GotFocus()
    Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Proyec)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_@.")
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_numTel)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " ")
   End If
End Sub

Private Sub txt_numCel_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)


   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
      ' ************************ 15012020 INICIO BY RAT
      Dim val As Integer
    
      val = validacionclientexproblemas(txt_NumDoc.Text, "")
     'val = 0
      
      '**** esto se va definir
      Dim var As Integer
      var = 110
      '*** esto se va definir

             g_str_Parame = ""
             g_str_Parame = "USP_CRE_INSPEK ("
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
             g_str_Parame = g_str_Parame & "'" & var & "', "
             g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
             g_str_Parame = g_str_Parame & "'" & Trim(modgen_g_str_rptwebservice) & "', "
             g_str_Parame = g_str_Parame & "'" & Trim(txt_Nombre) & "', "
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
             g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"

            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            Debug.Print ("invalido");
            Else
            Debug.Print ("valido");
            End If


         'Enviando Correo Electrónico
              modgen_g_str_Mail_Asunto = "REPORTE INSPEKTOR (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
              'modgen_g_str_Mail_Asunto = "REPORTE INSPEKTOR  "Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & "
              modgen_g_str_Mail_Mensaj = ""
              modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "DOC CLIENTE      : " & txt_NumDoc & Chr(13)
              modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "MENSAJE      : " & modgen_g_str_rptwebservice & Chr(13)
              modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
              modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
              modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)

              'MsgBox (modgen_g_str_Mail_Mensaj)

              If val <> 0 Then
              Debug.Print "CLIENTE NO TIENE PROBLEMAS"
              Else
              Debug.Print "CLIENTE  TIENE PROBLEMAS"
              MsgBox (modgen_g_str_rptwebservice)
              Call fs_Envia_CorreoInspektor(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
              Call fs_Limpia
              Call fs_Activa(False)
              cmb_TipDoc.Enabled = True
              txt_NumDoc.Enabled = True
              cmd_Buscar.Enabled = True
              cmd_GrabarProsp.Enabled = False
              Exit Sub
            End If
       ' ************************ 15012020 FIN BY RAT

   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)

'fs_Envia_CorreoEle
End If
   
End Sub

Private Sub txt_numTel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call gs_SetFocus(txt_numCel)
    Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
    End If
End Sub

Private Sub txt_Promo_GotFocus()
   Call gs_SelecTodo(txt_Promo)
End Sub

Private Sub txt_Promo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Constru)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_CadOri & "- ")
   End If
End Sub

Private Sub txt_Proyec_GotFocus()
   Call gs_SelecTodo(txt_Proyec)
End Sub

Private Sub txt_Proyec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Promo)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_CadOri & "- ")
   End If
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumDoc)
   End If
End Sub
