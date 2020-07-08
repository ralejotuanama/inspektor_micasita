VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_49 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5430
   Icon            =   "AteCli_frm_575.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5445
      _Version        =   65536
      _ExtentX        =   9604
      _ExtentY        =   6376
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
         TabIndex        =   9
         Top             =   30
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
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
            Height          =   255
            Left            =   660
            TabIndex        =   10
            Top             =   30
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes Desembolsadas"
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
         Begin Threed.SSPanel ssp_TipCon 
            Height          =   255
            Left            =   660
            TabIndex        =   11
            Top             =   300
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7064
            _ExtentY        =   450
            _StockProps     =   15
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
            Picture         =   "AteCli_frm_575.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   12
         Top             =   750
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4740
            Picture         =   "AteCli_frm_575.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   60
            Picture         =   "AteCli_frm_575.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   660
            Picture         =   "AteCli_frm_575.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1230
            Top             =   30
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   13
         Top             =   2140
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
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
         Begin VB.CheckBox chk_TipCon 
            Caption         =   "Todos"
            Height          =   255
            Left            =   1350
            TabIndex        =   2
            Top             =   420
            Width           =   3285
         End
         Begin VB.ComboBox cmb_TipCon 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   3975
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1350
            TabIndex        =   3
            Top             =   720
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1350
            TabIndex        =   4
            Top             =   1050
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin VB.Label lbl_TipCon 
            Caption         =   "Tipo de Consulta:"
            Height          =   465
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   225
            Left            =   60
            TabIndex        =   14
            Top             =   1050
            Width           =   1035
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   675
         Left            =   30
         TabIndex        =   17
         Top             =   1440
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
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
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   3975
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Reporte:"
            Height          =   285
            Left            =   60
            TabIndex        =   18
            Top             =   210
            Width           =   1260
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_49"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()   As moddat_tpo_Genera
Dim l_arr_ConHip()   As moddat_tpo_Genera

Private Sub chk_TipCon_Click()
   If chk_TipCon.Value = 1 Then
      cmb_TipCon.ListIndex = -1
      cmb_TipCon.Enabled = False
      Call gs_SetFocus(cmd_Imprim)
   ElseIf chk_TipCon.Value = 0 Then
      cmb_TipCon.Enabled = True
      Call gs_SetFocus(cmb_TipCon)
   End If
End Sub

Private Sub cmb_TipCon_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_TipCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipCon_Click
   End If
End Sub

Private Sub cmb_TipRep_Click()
   Call fs_Limpia
   If cmb_TipRep.ListIndex <> -1 Then
      If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
         Call moddat_gs_Carga_Produc(cmb_TipCon, l_arr_Produc, 4)
         chk_TipCon.Caption = "Todos los Productos"
         lbl_TipCon.Caption = "Producto:"
         ssp_TipCon.Caption = "Por Producto"
      Else
         Call moddat_gs_Carga_EjecMC(cmb_TipCon, l_arr_ConHip, 121)
         chk_TipCon.Caption = "Todos los Consejeros Hipotecarios"
         lbl_TipCon.Caption = "Consejero Hipotecario:"
         ssp_TipCon.Caption = "Por Consejero Hipotecario"
      End If
   End If
   Call gs_SetFocus(cmb_TipCon)
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipRep_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   If chk_TipCon.Value = 0 Then
      If cmb_TipCon.ListIndex = -1 Then
         If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
            MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         End If
         Call gs_SetFocus(cmb_TipCon)
         Exit Sub
      End If
   End If
   
   If cmb_TipCon.ListIndex <> -1 Then
      If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
         MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIni)
         Exit Sub
      End If
   End If
   
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   'Confirmación
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      Call fs_GenExc_TipPro
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      Call fs_GenExc_ConHip
   End If
End Sub

Private Sub cmd_Imprim_Click()
   'Validación
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   If chk_TipCon.Value = 0 Then
      If cmb_TipCon.ListIndex = -1 Then
         If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
            MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         End If
         Call gs_SetFocus(cmb_TipCon)
         Exit Sub
      End If
   End If
   
   If cmb_TipCon.ListIndex <> -1 Then
      If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
         MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIni)
         Exit Sub
      End If
   End If
   
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
    'Confirmación
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
       
   'Se modifica el puntero para un estado de espera
   Screen.MousePointer = 11
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      Call fs_GenImp_TipPro
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      Call fs_GenImp_ConHip
   End If
   
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Limpia
   cmb_TipRep.AddItem "POR PRODUCTO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
   If modgen_g_int_TipUsu <> 20121 Then
      cmb_TipRep.AddItem "POR CONSEJERO HIPOTECARIO"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2
   End If
   cmb_TipRep.ListIndex = -1
   
   Call gs_CentraForm(Me)
   Call gs_SetFocus(cmb_TipCon)
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenImp_ConHip()
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_SOLMAE"
      
   'Se selecciona la formula
   crp_Imprim.SelectionFormula = ""
   
   'Se Filtra por el tipo de producto escogido en el formulario
   If chk_TipCon.Value = 0 Then
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_SOLMAE.SOLMAE_CONHIP} = '" & l_arr_ConHip(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   'Se realiza la validación para codigo de instancia y fechas
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_SOLMAE.SOLMAE_SITUAC} = 2 AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_HIPMAE.HIPMAE_FECDES} >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_HIPMAE.HIPMAE_FECDES} <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
           
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_14.RPT"
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenImp_TipPro()
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      
   crp_Imprim.DataFiles(0) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_SOLMAE"
   
   'Se selecciona la formula
   crp_Imprim.SelectionFormula = ""
   
   'Se Filtra por el tipo de producto escogido en el formulario
   If chk_TipCon.Value = 0 Then
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_SOLMAE.SOLMAE_CODPRD} = '" & l_arr_Produc(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   'Se realiza la validación para codigo de instancia y fechas
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_SOLMAE.SOLMAE_SITUAC} = 2 AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_HIPMAE.HIPMAE_CONHIP} = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_HIPMAE.HIPMAE_FECDES} >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_HIPMAE.HIPMAE_FECDES} <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
        
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_13.RPT"
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenExc_ConHip()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT D.PRODUC_DESCRI, A.SOLMAE_NUMERO, A.SOLMAE_TITTDO, A.SOLMAE_TITNDO, A.SOLMAE_TIPMON, SOLMAE_CONHIP, "
   g_str_Parame = g_str_Parame & "        (TRIM(B.DATGEN_APEPAT) || ' ' || TRIM(B.DATGEN_APEMAT) || ' ' ||TRIM(B.DATGEN_NOMBRE)) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "        C.PARDES_DESCRI AS SEDE, A.SOLMAE_FECSOL, E.HIPMAE_FECDES, E.HIPMAE_NUMOPE, A.SOLMAE_TASINT, "
   g_str_Parame = g_str_Parame & "        E.HIPMAE_PLAANO , E.HIPMAE_PERGRA, A.SOLMAE_COMVTA_SOL, A.SOLMAE_COMVTA_DOL, "
   g_str_Parame = g_str_Parame & "        A.SOLMAE_APOPRO_SOL, A.SOLMAE_APOPRO_DOL, A.SOLMAE_MTOPRE_MPR "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES C ON C.PARDES_CODITE = A.SOLMAE_ATECOM AND C.PARDES_CODGRP = 518 "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC D ON D.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_HIPMAE E ON E.HIPMAE_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "  WHERE "
   
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
      
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC IN (2, 6, 9) AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CONHIP ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "CONSEJ. HIPOT."
      .Cells(1, 3) = "SEDE"
      .Cells(1, 4) = "PRODUCTO"
      .Cells(1, 5) = "SOLICITUD"
      .Cells(1, 6) = "DOC. IDENTIDAD"
      .Cells(1, 7) = "NOMBRE CLIENTE"
      .Cells(1, 8) = "F. SOLICITUD"
      .Cells(1, 9) = "F. DESEMBOLSO"
      .Cells(1, 10) = "OPERACION"
      .Cells(1, 11) = "TASA"
      .Cells(1, 12) = "PLAZO"
      .Cells(1, 13) = "PERIODO GRACIA"
      .Cells(1, 14) = "TIP. DE MONEDA"
      .Cells(1, 15) = "V. INMUEBLE S/."
      .Cells(1, 16) = "V. INMUEBLE US$."
      .Cells(1, 17) = "PORC. INICIAL"
      .Cells(1, 18) = "MTO. CREDITO S/."
      .Cells(1, 19) = "MTO. CREDITO US$."
      .Cells(1, 20) = "V. ASEGURABLE INM."
      
      .Range(.Cells(1, 1), .Cells(1, 20)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 20)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 19
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 21
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 33
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 40
      .Columns("H").ColumnWidth = 12
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 14
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 15
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 7
      .Columns("L").ColumnWidth = 6
      .Columns("M").ColumnWidth = 15
      .Columns("N").ColumnWidth = 21
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 17
      .Columns("P").ColumnWidth = 17
      .Columns("Q").ColumnWidth = 16
      .Columns("R").ColumnWidth = 18
      .Columns("S").ColumnWidth = 18
      .Columns("T").ColumnWidth = 20
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!SEDE) 'Trim(g_rst_Princi!PARDES_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!PRODUC_DESCRI) 'moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!CLIENTE) 'moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = CStr(g_rst_Princi!SOLMAE_TASINT) + "%"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = g_rst_Princi!HIPMAE_PLAANO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = g_rst_Princi!HIPMAE_PERGRA
 
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = "SOLES"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = "DOLARES AMERICANOS"
      End If
 
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
      End If
      
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
      End If
      
      If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Or g_rst_Princi!SOLMAE_COMVTA_DOL > 0 Then
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_SOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00") * 100) + "%"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_DOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00") * 100) + "%"
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = CStr(0) + "%"
      End If
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = 0
      End If
      
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = 0
      End If
       
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Format(fs_Buscar_ValorAseg(Trim(g_rst_Princi!SOLMAE_NUMERO)), "###,###,##0.00")
                      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_TipPro()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer

   g_str_Parame = "SELECT TRIM(D.PRODUC_DESCRI) AS PRODUCTO, "
   g_str_Parame = g_str_Parame & "       TRIM(G.SUBPRD_DESCRI) AS SUB_PRODUCTO, "
   g_str_Parame = g_str_Parame & "       A.SOLMAE_NUMERO AS SOLICITUD, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_TIPDOC)||'-'||TRIM(C.DATGEN_NUMDOC) AS DOCIDENTIDAD, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT)||' '||TRIM(C.DATGEN_NOMBRE) AS NOMCLIENTE, "
   g_str_Parame = g_str_Parame & "       A.SOLMAE_FECSOL AS FECHA_SOL, "
   g_str_Parame = g_str_Parame & "       B.HIPMAE_FECDES AS FECHA_DES, "
   g_str_Parame = g_str_Parame & "       A.SOLMAE_CONHIP AS CONSEJERO, "
   g_str_Parame = g_str_Parame & "       TRIM(E.PARDES_DESCRI) AS SEDE, "
   g_str_Parame = g_str_Parame & "       B.HIPMAE_NUMOPE AS OPERACION, "
   g_str_Parame = g_str_Parame & "       B.HIPMAE_TASINT AS TASA, "
   g_str_Parame = g_str_Parame & "       B.HIPMAE_PLAANO AS PLAZO, "
   g_str_Parame = g_str_Parame & "       B.HIPMAE_PERGRA AS PER_GRACIA, "
   g_str_Parame = g_str_Parame & "       A.SOLMAE_TIPMON AS MONEDA, "
   g_str_Parame = g_str_Parame & "       DECODE(A.SOLMAE_TIPMON, 1, SOLMAE_COMVTA_SOL, SOLMAE_COMVTA_DOL) AS VAL_INMUEBLE, "
   g_str_Parame = g_str_Parame & "       ROUND(NVL(DECODE(A.SOLMAE_TIPMON, 1, SOLMAE_APOPRO_SOL/SOLMAE_COMVTA_SOL, SOLMAE_APOPRO_DOL/SOLMAE_COMVTA_DOL)*100,0),2) AS PORC_INICIAL, "
   g_str_Parame = g_str_Parame & "       A.SOLMAE_MTOPRE_MPR AS MTO_CREDITO, "
   g_str_Parame = g_str_Parame & "       F.EVATAS_SUMASE_INM + F.EVATAS_SUMASE_ES1 + F.EVATAS_SUMASE_ES2 + F.EVATAS_SUMASE_DEP AS VAL_ASEGURABLE, "
   g_str_Parame = g_str_Parame & "       NVL(TRIM(J.PARDES_DESCRI),'-') AS PRY_MICASITA, "
   g_str_Parame = g_str_Parame & "       TRIM(NVL(DECODE(H.SOLINM_PRYCOD, 1, H.SOLINM_PRYNOM, DECODE(H.SOLINM_PRYCOD, NULL, H.SOLINM_PRYNOM, I.DATGEN_TITULO)),'-') ) AS PROYECTO, "
   g_str_Parame = g_str_Parame & "       TRIM(H.SOLINM_TIPDOC_PRO ||'-'|| H.SOLINM_NUMDOC_PRO) AS DOCPROMOTOR, "
   g_str_Parame = g_str_Parame & "       NVL(CASE WHEN H.SOLINM_TIPDOC_PRO = 7 THEN TRIM(K.DATGEN_RAZSOC) ELSE  TRIM(H.SOLINM_RAZSOC_PRO) END, '-') AS NOM_PROMOTOR, "
   g_str_Parame = g_str_Parame & "       NVL(CASE WHEN SOLINM_TIPDOC_CON= 0 THEN '-' ELSE TRIM(SOLINM_TIPDOC_CON ||'-'|| SOLINM_NUMDOC_CON) END, '-') AS DOCCONSTRUCTOR, "
   g_str_Parame = g_str_Parame & "       NVL(CASE WHEN H.SOLINM_TIPDOC_CON = 7 THEN TRIM(L.DATGEN_RAZSOC) ELSE  TRIM(H.SOLINM_RAZSOC_CON) END, '-') AS NOM_CONSTRUCTOR "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "INNER JOIN CRE_HIPMAE B "
   g_str_Parame = g_str_Parame & "    ON B.HIPMAE_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "INNER JOIN CLI_DATGEN C "
   g_str_Parame = g_str_Parame & "    ON C.DATGEN_TIPDOC = A.SOLMAE_TITTDO AND C.DATGEN_NUMDOC = A.SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & "INNER JOIN CRE_PRODUC D "
   g_str_Parame = g_str_Parame & "    ON D.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "INNER JOIN MNT_PARDES E "
   g_str_Parame = g_str_Parame & "    ON E.PARDES_CODGRP = 518 AND E.PARDES_CODITE = A.SOLMAE_ATECOM "
   g_str_Parame = g_str_Parame & "INNER JOIN TRA_EVATAS F "
   g_str_Parame = g_str_Parame & "    ON F.EVATAS_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "INNER JOIN CRE_SUBPRD G "
   g_str_Parame = g_str_Parame & "    ON G.SUBPRD_CODPRD = A.SOLMAE_CODPRD AND G.SUBPRD_CODSUB = A.SOLMAE_CODSUB "
   g_str_Parame = g_str_Parame & "LEFT JOIN CRE_SOLINM H "
   g_str_Parame = g_str_Parame & "    ON H.SOLINM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "LEFT JOIN PRY_DATGEN I "
   g_str_Parame = g_str_Parame & "    ON I.DATGEN_CODIGO = H.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "LEFT JOIN MNT_PARDES J "
   g_str_Parame = g_str_Parame & "    ON J.PARDES_CODGRP = 214 AND J.PARDES_CODITE = H.SOLINM_PRYMCS "
   g_str_Parame = g_str_Parame & "LEFT JOIN EMP_DATGEN K "
   g_str_Parame = g_str_Parame & "    ON K.DATGEN_EMPTDO = H.SOLINM_TIPDOC_PRO AND K.DATGEN_EMPNDO = H.SOLINM_NUMDOC_PRO "
   g_str_Parame = g_str_Parame & "LEFT JOIN EMP_DATGEN L "
   g_str_Parame = g_str_Parame & "    ON L.DATGEN_EMPTDO = H.SOLINM_TIPDOC_CON AND L.DATGEN_EMPNDO = H.SOLINM_NUMDOC_CON "
   g_str_Parame = g_str_Parame & " WHERE B.HIPMAE_SITUAC IN (2, 6, 9) "
   g_str_Parame = g_str_Parame & "   AND B.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND B.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   
   'Restricción por producto
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND A.SOLMAE_CODPRD = '" & l_arr_Produc(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' "
   End If
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "   AND A.SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CODPRD ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "PRODUCTO"
      .Cells(1, 3) = "SUB-PRODUCTO"
      .Cells(1, 4) = "SOLICITUD"
      .Cells(1, 5) = "DOC. IDENTIDAD"
      .Cells(1, 6) = "NOMBRE CLIENTE"
      .Cells(1, 7) = "F. SOLICITUD"
      .Cells(1, 8) = "F. DESEMBOLSO"
      .Cells(1, 9) = "CONSEJ. HIPOT."
      .Cells(1, 10) = "SEDE"
      .Cells(1, 11) = "OPERACION"
      .Cells(1, 12) = "TASA"
      .Cells(1, 13) = "PLAZO"
      .Cells(1, 14) = "PERIODO GRACIA"
      .Cells(1, 15) = "TIP. MONEDA"
      .Cells(1, 16) = "V. INMUEBLE"
      .Cells(1, 17) = "PORC. INICIAL"
      .Cells(1, 18) = "MTO. CREDITO"
      .Cells(1, 19) = "V. ASEGURABLE INM."
      .Cells(1, 20) = "PRY MICASITA"
      .Cells(1, 21) = "PROYECTO"
      .Cells(1, 22) = "DOI PROMOTOR"
      .Cells(1, 23) = "NOMBRE PROMOTOR"
      .Cells(1, 24) = "DOI CONSTRUCTOR"
      .Cells(1, 25) = "NOMBRE CONSTRUCTOR"
               
      .Range(.Cells(1, 1), .Cells(1, 25)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 25)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 6
      .Columns("B").ColumnWidth = 40
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 60
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 16
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 16
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 40
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 15
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 20
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 20
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 16
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 10
      .Columns("M").ColumnWidth = 10
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").ColumnWidth = 16
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 20
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 15
      .Columns("Q").ColumnWidth = 15
      .Columns("R").ColumnWidth = 18
      .Columns("S").ColumnWidth = 20
      .Columns("T").ColumnWidth = 14
      .Columns("T").HorizontalAlignment = xlHAlignCenter
      .Columns("U").ColumnWidth = 50
      .Columns("U").HorizontalAlignment = xlHAlignCenter
      .Columns("V").ColumnWidth = 18
      .Columns("V").HorizontalAlignment = xlHAlignCenter
      .Columns("W").ColumnWidth = 50
      .Columns("W").HorizontalAlignment = xlHAlignCenter
      .Columns("X").ColumnWidth = 18
      .Columns("X").HorizontalAlignment = xlHAlignCenter
      .Columns("Y").ColumnWidth = 50
      .Columns("Y").HorizontalAlignment = xlHAlignCenter
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!SUB_PRODUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumSol(g_rst_Princi!SOLICITUD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!DOCIDENTIDAD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!NOMCLIENTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_SOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_DES)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!CONSEJERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!SEDE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = gf_Formato_NumOpe(Trim(g_rst_Princi!OPERACION))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CStr(g_rst_Princi!TASA) + "%"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = g_rst_Princi!PLAZO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = g_rst_Princi!PER_GRACIA
      If g_rst_Princi!MONEDA = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = "SOLES"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = "DOLARES AMERICANOS"
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!VAL_INMUEBLE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = CStr(Format(g_rst_Princi!PORC_INICIAL, "###,###,##0.00")) + "%"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(g_rst_Princi!MTO_CREDITO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(g_rst_Princi!VAL_ASEGURABLE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Trim(g_rst_Princi!PRY_MICASITA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Trim(g_rst_Princi!PROYECTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Trim(g_rst_Princi!DOCPROMOTOR)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Trim(g_rst_Princi!NOM_PROMOTOR)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Trim(g_rst_Princi!DOCCONSTRUCTOR)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Trim(g_rst_Princi!NOM_CONSTRUCTOR)
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Function fs_Buscar_ValorAseg(ByVal numero As String) As Double
   Dim g_str_cadena  As String
   fs_Buscar_ValorAseg = 0
   
   g_str_cadena = ""
   g_str_cadena = "SELECT ( EVATAS_SUMASE_INM + EVATAS_SUMASE_ES1 + EVATAS_SUMASE_ES2 + EVATAS_SUMASE_DEP ) as TOTAL "
   g_str_cadena = g_str_cadena & "FROM TRA_EVATAS WHERE "
   g_str_cadena = g_str_cadena & "EVATAS_NUMSOL = '" & numero & "' "

   If Not gf_EjecutaSQL(g_str_cadena, g_rst_GenAux, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      fs_Buscar_ValorAseg = gf_FormatoNumero(g_rst_GenAux!Total, 12, 2)
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Function

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Imprim)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub fs_Limpia()
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
   cmb_TipCon.Clear
   chk_TipCon.Value = 0
   Call gs_SetFocus(cmb_TipCon)
End Sub
