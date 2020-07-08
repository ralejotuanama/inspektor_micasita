VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_48 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5430
   Icon            =   "AteCli_frm_574.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3585
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5445
      _Version        =   65536
      _ExtentX        =   9604
      _ExtentY        =   6324
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
            Left            =   630
            TabIndex        =   10
            Top             =   30
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes Rechazadas"
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
            Left            =   630
            TabIndex        =   11
            Top             =   300
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
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
            Picture         =   "AteCli_frm_574.frx":000C
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_574.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4740
            Picture         =   "AteCli_frm_574.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_574.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1740
            Top             =   90
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
         Height          =   1395
         Left            =   30
         TabIndex        =   13
         Top             =   2140
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
         _ExtentY        =   2461
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
         Begin VB.ComboBox cmb_TipCon 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   3915
         End
         Begin VB.CheckBox chk_TipCon 
            Caption         =   "Todos"
            Height          =   255
            Left            =   1350
            TabIndex        =   2
            Top             =   390
            Width           =   3825
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1350
            TabIndex        =   3
            Top             =   660
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
            Top             =   1020
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
            Height          =   435
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   225
            Left            =   60
            TabIndex        =   15
            Top             =   1020
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   255
            Left            =   60
            TabIndex        =   14
            Top             =   660
            Width           =   1065
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
            Width           =   3915
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
Attribute VB_Name = "frm_RptSol_48"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_ConHip()      As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

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
   If chk_TipCon.Value <> 0 Then
      If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
         MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIni)
         Exit Sub
      End If
   End If
      
   'Confirmación
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Proceso
   Screen.MousePointer = 11
  
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      Call fs_GenImp_TipPro
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      Call fs_GenImp_ConHip
   End If
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
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
   If chk_TipCon.Value <> 0 Then
      If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
         MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIni)
         Exit Sub
      End If
   End If
      
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      Call fs_GenExc_TipPro
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      Call fs_GenExc_ConHip
   End If
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

Private Sub fs_Limpia()
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
   cmb_TipCon.Clear
   chk_TipCon.Value = 0
   Call gs_SetFocus(cmb_TipCon)
End Sub

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

Private Function ff_Observ(ByVal p_NumSol As String) As String

   g_str_Parame = "SELECT SEGDET_OBSERV FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "(SEGDET_CODOCU = 13 OR SEGDET_CODOCU = 14 )"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      If g_rst_Listas!SEGDET_OBSERV <> "" Then
         Do While Not g_rst_Listas.EOF
            ff_Observ = g_rst_Listas!SEGDET_OBSERV
            g_rst_Listas.MoveNext
         Loop
      End If
   End If
    
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_GenImp_ConHip()
Dim r_str_CodRec     As String
Dim r_str_Observ     As String

   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
   
   'Eliminamos el contenido de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_MOTREC WHERE "
   g_str_Parame = g_str_Parame & "MOTREC_NOMRPT = 'ATE_RPTSOL_16.RPT' AND "
   g_str_Parame = g_str_Parame & "MOTREC_TERCRE = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DISTINCT SOLMAE_NUMERO, SOLMAE_MOTREC, B.PARDES_DESCRI AS MOTIVO_RECHAZO, C.SEGDET_OBSERV  "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES B ON B.PARDES_CODITE = A.SOLMAE_MOTREC AND B.PARDES_CODGRP = '003' "
   g_str_Parame = g_str_Parame & "         LEFT JOIN TRA_SEGDET C ON C.SEGDET_NUMSOL = A.SOLMAE_NUMERO AND (C.SEGDET_CODOCU = 13 OR C.SEGDET_CODOCU = 14) "
   g_str_Parame = g_str_Parame & "  WHERE "
   
   'Si no escogio todos los Productos
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_FECREC >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECREC <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         'Para obtener Descripción de Motivos de Rechazo
         r_str_CodRec = Trim(g_rst_Princi!MOTIVO_RECHAZO) 'moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
         r_str_Observ = IIf(IsNull(Trim(g_rst_Princi!SEGDET_OBSERV)), "", Trim(g_rst_Princi!SEGDET_OBSERV)) 'ff_Observ(g_rst_Princi!SOLMAE_NUMERO)
         
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_MOTREC("
         g_str_Parame = g_str_Parame & "MOTREC_NOMRPT, "
         g_str_Parame = g_str_Parame & "MOTREC_FECCRE, "
         g_str_Parame = g_str_Parame & "MOTREC_HORCRE, "
         g_str_Parame = g_str_Parame & "MOTREC_TERCRE, "
         g_str_Parame = g_str_Parame & "MOTREC_NUMSOL, "
         g_str_Parame = g_str_Parame & "MOTREC_CODREC, "
         g_str_Parame = g_str_Parame & "MOTREC_OBSERV) "
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_16.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_CodRec & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_Observ & "') "
          
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Else
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CLI_DATGEN"
   crp_Imprim.DataFiles(2) = "CRE_PRODUC"
   crp_Imprim.DataFiles(3) = "RPT_MOTREC"
     
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_16.RPT"
   crp_Imprim.SelectionFormula = "{RPT_MOTREC.MOTREC_NOMRPT} = 'ATE_RPTSOL_16.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_MOTREC.MOTREC_TERCRE} = '" & modgen_g_str_NombPC & "' "
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenImp_TipPro()
Dim r_str_CodRec     As String
Dim r_str_Observ     As String

   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
   
   'Eliminamos el contenido de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_MOTREC WHERE "
   g_str_Parame = g_str_Parame & "MOTREC_NOMRPT = 'ATE_RPTSOL_15.RPT' AND "
   g_str_Parame = g_str_Parame & "MOTREC_TERCRE = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DISTINCT SOLMAE_NUMERO, B.PARDES_DESCRI AS MOTIVO_RECHAZO , C.SEGDET_OBSERV "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES B ON B.PARDES_CODITE = A.SOLMAE_MOTREC AND B.PARDES_CODGRP = '003' "
   g_str_Parame = g_str_Parame & "         LEFT JOIN TRA_SEGDET C ON C.SEGDET_NUMSOL = A.SOLMAE_NUMERO AND (C.SEGDET_CODOCU = 13 OR C.SEGDET_CODOCU = 14) "
   g_str_Parame = g_str_Parame & "  WHERE "
   
   'Si no escogio todos los Productos
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   g_str_Parame = g_str_Parame & "SOLMAE_FECREC >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECREC <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
                  
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_CodRec = Trim(g_rst_Princi!MOTIVO_RECHAZO) 'moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
         r_str_Observ = IIf(IsNull(Trim(g_rst_Princi!SEGDET_OBSERV)), "", g_rst_Princi!SEGDET_OBSERV) 'ff_Observ(g_rst_Princi!SOLMAE_NUMERO) '
                    
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_MOTREC("
         g_str_Parame = g_str_Parame & "MOTREC_NOMRPT, "
         g_str_Parame = g_str_Parame & "MOTREC_FECCRE, "
         g_str_Parame = g_str_Parame & "MOTREC_HORCRE, "
         g_str_Parame = g_str_Parame & "MOTREC_TERCRE, "
         g_str_Parame = g_str_Parame & "MOTREC_NUMSOL, "
         g_str_Parame = g_str_Parame & "MOTREC_CODREC, "
         g_str_Parame = g_str_Parame & "MOTREC_OBSERV) "
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_15.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_CodRec & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_Observ & "') "
                  
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
            
         g_rst_Princi.MoveNext
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Else
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CLI_DATGEN"
   crp_Imprim.DataFiles(2) = "CRE_PRODUC"
   crp_Imprim.DataFiles(3) = "RPT_MOTREC"
   crp_Imprim.SelectionFormula = "{RPT_MOTREC.MOTREC_NOMRPT} = 'ATE_RPTSOL_15.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_MOTREC.MOTREC_TERCRE} = '" & modgen_g_str_NombPC & "' "
   
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_15.RPT"
      
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenExc_ConHip()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT C.PRODUC_DESCRI, A.SOLMAE_NUMERO, A.SOLMAE_TITTDO, A.SOLMAE_TITNDO, A.SOLMAE_CONHIP, "
   g_str_Parame = g_str_Parame & "        (TRIM(B.DATGEN_APEPAT) || ' ' || TRIM(B.DATGEN_APEMAT) || ' ' ||TRIM(B.DATGEN_NOMBRE)) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "        B.DATGEN_DIRELE, B.DATGEN_NUMCEL, B.DATGEN_TELEFO, A.SOLMAE_FECSOL, A.SOLMAE_FECREC, "
   g_str_Parame = g_str_Parame & "        A.SOLMAE_MOTREC, D.SEGDET_OBSERV, E.PARDES_DESCRI AS MOTIVO_RECHAZO "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_SEGDET D ON D.SEGDET_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES E ON E.PARDES_CODITE = A.SOLMAE_MOTREC AND E.PARDES_CODGRP = '003'"
   g_str_Parame = g_str_Parame & "  WHERE "
   
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
      
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 AND "
   g_str_Parame = g_str_Parame & "(SEGDET_CODOCU=13 OR SEGDET_CODOCU=14) AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECREC >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECREC <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
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
      .Cells(1, 3) = "PRODUCTO"
      .Cells(1, 4) = "SOLICITUD"
      .Cells(1, 5) = "DOC. IDENTIDAD"
      .Cells(1, 6) = "NOMBRE CLIENTE"
      .Cells(1, 7) = "CORREO ELECTRÓNICO"
      .Cells(1, 8) = "TELÉFONO"
      .Cells(1, 9) = "NOMBRE DEL PROYECTO"
      .Cells(1, 10) = "F. INGRESO"
      .Cells(1, 11) = "F. RECHAZO"
      .Cells(1, 12) = "MOTIVO DE RECHAZO"
      .Cells(1, 13) = "OBSERVACIONES"
         
      .Range(.Cells(1, 1), .Cells(1, 10)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 10)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 15
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 35
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 40
      .Columns("G").ColumnWidth = 40
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 25
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 45
      .Columns("J").ColumnWidth = 15
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 15
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 46
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 100
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!PRODUC_DESCRI) 'moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!CLIENTE) 'moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!DatGen_DirEle)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(IIf(IsNull(g_rst_Princi!DATGEN_NUMCEL), "", g_rst_Princi!DATGEN_NUMCEL)) & IIf(IsNull(g_rst_Princi!DATGEN_NUMCEL), "", IIf(IsNull(g_rst_Princi!DatGen_Telefo), "", "; ")) & Trim(IIf(IsNull(g_rst_Princi!DatGen_Telefo), "", g_rst_Princi!DatGen_Telefo))
      
      'Obteniendo Información del Inmueble
      Call moddat_gs_Consulta_DatInm(g_rst_Princi!SOLMAE_NUMERO, "", "", r_str_CodPry, r_str_NomPry, r_str_CodBco)
      
      If Len(Trim(r_str_CodPry)) > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = moddat_gf_Consulta_NomPry(r_str_CodPry)
      ElseIf Len(Trim(r_str_NomPry)) > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = r_str_NomPry '& " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!MOTIVO_RECHAZO) 'moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Trim(g_rst_Princi!SEGDET_OBSERV)
       
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   r_obj_Excel.ActiveSheet.Cells(1, 8).HorizontalAlignment = xlHAlignCenter
    
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_TipPro()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String
   
'   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, CLI_DATGEN B, TRA_SEGDET C WHERE "
'   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = DATGEN_TIPDOC AND "
'   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = DATGEN_NUMDOC AND "
'   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SEGDET_NUMSOL AND "

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT C.PRODUC_DESCRI, A.SOLMAE_NUMERO, A.SOLMAE_TITTDO, A.SOLMAE_TITNDO, A.SOLMAE_CONHIP, "
   g_str_Parame = g_str_Parame & "        (TRIM(B.DATGEN_APEPAT) || ' ' || TRIM(B.DATGEN_APEMAT) || ' ' ||TRIM(B.DATGEN_NOMBRE)) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "        B.DATGEN_DIRELE, B.DATGEN_NUMCEL, B.DATGEN_TELEFO, A.SOLMAE_FECSOL, A.SOLMAE_FECREC, "
   g_str_Parame = g_str_Parame & "        A.SOLMAE_MOTREC, D.SEGDET_OBSERV, E.PARDES_DESCRI AS MOTIVO_RECHAZO "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_SEGDET D ON D.SEGDET_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES E ON E.PARDES_CODITE = A.SOLMAE_MOTREC AND E.PARDES_CODGRP = '003'"
   g_str_Parame = g_str_Parame & "  WHERE "
      
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
      
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "(SEGDET_CODOCU = 13 OR SEGDET_CODOCU = 14) AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECREC >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECREC <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CODPRD ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC "
    
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
      .Cells(1, 3) = "SOLICITUD"
      .Cells(1, 4) = "DOC. IDENTIDAD"
      .Cells(1, 5) = "NOMBRE CLIENTE"
      .Cells(1, 6) = "CORREO ELECTRÓNICO"
      .Cells(1, 7) = "TELÉFONO"
      .Cells(1, 8) = "NOMBRE DEL PROYECTO"
      .Cells(1, 9) = "F. INGRESO"
      .Cells(1, 10) = "F. RECHAZO"
      .Cells(1, 11) = "CONSEJ. HIPOT."
      .Cells(1, 12) = "MOTIVO DE RECHAZO"
      .Cells(1, 13) = "OBSERVACIONES"
         
      .Range(.Cells(1, 1), .Cells(1, 10)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 10)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 35
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 40
      .Columns("F").ColumnWidth = 40
      .Columns("G").ColumnWidth = 25
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 45
      .Columns("I").ColumnWidth = 15
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 15
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 15
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 46
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 100
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI) 'moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!CLIENTE) 'moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!DatGen_DirEle)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(IIf(IsNull(g_rst_Princi!DATGEN_NUMCEL), "", g_rst_Princi!DATGEN_NUMCEL)) & IIf(IsNull(g_rst_Princi!DATGEN_NUMCEL), "", IIf(IsNull(g_rst_Princi!DatGen_Telefo), "", "; ")) & Trim(IIf(IsNull(g_rst_Princi!DatGen_Telefo), "", g_rst_Princi!DatGen_Telefo))

      'Obteniendo Información del Inmueble
      Call moddat_gs_Consulta_DatInm(g_rst_Princi!SOLMAE_NUMERO, "", "", r_str_CodPry, r_str_NomPry, r_str_CodBco)
      
      If Len(Trim(r_str_CodPry)) > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = moddat_gf_Consulta_NomPry(r_str_CodPry)
      ElseIf Len(Trim(r_str_NomPry)) > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = r_str_NomPry '& " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!MOTIVO_RECHAZO) 'moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Trim(g_rst_Princi!SEGDET_OBSERV)
       
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   r_obj_Excel.ActiveSheet.Cells(1, 7).HorizontalAlignment = xlHAlignCenter
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

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
