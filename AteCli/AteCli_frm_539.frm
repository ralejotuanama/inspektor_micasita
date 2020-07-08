VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_33 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   2220
   ClientTop       =   4425
   ClientWidth     =   5415
   Icon            =   "AteCli_frm_539.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3285
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5445
      _Version        =   65536
      _ExtentX        =   9604
      _ExtentY        =   5794
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
            Height          =   315
            Left            =   690
            TabIndex        =   10
            Top             =   30
            Width           =   3405
            _Version        =   65536
            _ExtentX        =   6006
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes en Instancia"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Left            =   690
            TabIndex        =   11
            Top             =   300
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Por Consejero Hipotecario"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Picture         =   "AteCli_frm_539.frx":000C
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_539.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_539.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4740
            Picture         =   "AteCli_frm_539.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
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
         Height          =   1785
         Left            =   30
         TabIndex        =   13
         Top             =   1440
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
         _ExtentY        =   3149
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
         Begin VB.CheckBox chk_ConHip 
            Caption         =   "Todos los Consejero Hipotecario"
            Height          =   315
            Left            =   1230
            TabIndex        =   1
            Top             =   420
            Width           =   2685
         End
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   4065
         End
         Begin VB.CheckBox Chk_FecCon 
            Caption         =   "Todas las Solicitudes en Trámite"
            Height          =   315
            Left            =   1230
            TabIndex        =   4
            Top             =   1410
            Width           =   2685
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1230
            TabIndex        =   2
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
            Left            =   1230
            TabIndex        =   3
            Top             =   1080
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
         Begin VB.Label Label4 
            Caption         =   "Consejero Hipotecario:"
            Height          =   465
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1005
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   780
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   225
            Left            =   60
            TabIndex        =   14
            Top             =   1110
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_33"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_ConHip()      As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

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
   If Chk_FecCon.Value = 1 Then
      ipp_FecIni.Enabled = False
      ipp_FecFin.Enabled = False
      Call gs_SetFocus(cmd_Imprim)
   ElseIf Chk_FecCon.Value = 0 Then
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

Private Sub cmd_ExpExc_Click()
   If chk_ConHip.Value = 0 Then
      If cmb_ConHip.ListIndex = -1 Then
         MsgBox "Debe seleccionar a un Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ConHip)
         Exit Sub
      End If
   End If
   
   'Validación
   If chk_ConHip.Value = 0 Then
      If cmb_ConHip.ListIndex = -1 Then
         MsgBox "Debe seleccionar a un Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ConHip)
         Exit Sub
      End If
   End If
   If cmb_ConHip.ListIndex <> -1 Then
      If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
         MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIni)
         Exit Sub
      End If
   End If
   If Chk_FecCon.Value = 0 Then
      If chk_ConHip <> 0 Then
         If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
            MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_FecIni)
            Exit Sub
         End If
      End If
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   
   Call fs_GenExc
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_str_DesOcu     As String
   Dim r_str_SitSol     As String
   Dim r_str_CodIns     As String
   
   'Validación
   If chk_ConHip.Value = 0 Then
      If cmb_ConHip.ListIndex = -1 Then
         MsgBox "Debe seleccionar a un Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ConHip)
         Exit Sub
      End If
   End If
   
   If cmb_ConHip.ListIndex <> -1 Then
      If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
         MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIni)
         Exit Sub
      End If
   End If
   
   If Chk_FecCon.Value = 0 Then
      If chk_ConHip <> 0 Then
         If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
            MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_FecIni)
            Exit Sub
         End If
      End If
   End If
   
   'Confirmación
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   'Proceso
   Screen.MousePointer = 11
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   'Eliminamos el contenido de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT = 'ATE_RPTSOL_28.RPT' AND "
   g_str_Parame = g_str_Parame & "SOLTRA_TERCRE = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   'Si no escogio todos los Consejeros Hipotecarios
   If chk_ConHip.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If Chk_FecCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND "
   Else
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   End If
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL > 0 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_CodIns = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SOLMAE_CODINS))
         
         'Para obtener Descripción de Instancia Actual
         r_str_SitSol = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
      
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_SOLTRA("
         g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT, "
         g_str_Parame = g_str_Parame & "SOLTRA_FECCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_HORCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_TERCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_NUMSOL, "
         g_str_Parame = g_str_Parame & "SOLTRA_CODOCU, "
         g_str_Parame = g_str_Parame & "SOLTRA_SITIN1) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_28.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_CodIns & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_SitSol & "') "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = UCase(moddat_g_str_EntDat) & ".CRE_PRODUC"
   crp_Imprim.DataFiles(2) = UCase(moddat_g_str_EntDat) & ".CLI_DATGEN"
   crp_Imprim.DataFiles(3) = UCase(moddat_g_str_EntDat) & ".RPT_SOLTRA"
   
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_28.RPT"
   
   crp_Imprim.SelectionFormula = "{RPT_SOLTRA.SOLTRA_NOMRPT} = 'ATE_RPTSOL_28.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SOLTRA.SOLTRA_TERCRE} = '" & modgen_g_str_NombPC & "' "
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_ConHip)
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, CLI_DATGEN B WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = DATGEN_TIPDOC AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = DATGEN_NUMDOC AND "
      
   If chk_ConHip.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "' AND "
   End If
      
   If Chk_FecCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND "
   Else
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   End If
      
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL > 0 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CONHIP ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC, SOLMAE_FECSOL DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
    
   'Si no encuentra ninguna Solicitud
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
      .Cells(1, 7) = "F. SOLICITUD"
      .Cells(1, 8) = "INSTANCIA ACTUAL"
      .Cells(1, 9) = "SITUACION DE SOLICITUD"
      .Cells(1, 10) = "TIP. MONEDA"
      .Cells(1, 11) = "V. INMUEBLE S/."
      .Cells(1, 12) = "V. INMUEBLE US$."
      .Cells(1, 13) = "PORC. INICIAL"
      .Cells(1, 14) = "MTO. CREDITO S/."
      .Cells(1, 15) = "MTO. CREDITO US$."
     
   
      .Range(.Cells(1, 1), .Cells(1, 15)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 15)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 5
      
      .Columns("B").ColumnWidth = 15
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 31
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 40
            
      .Columns("G").ColumnWidth = 12
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 30
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 52
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 21
      .Columns("J").HorizontalAlignment = xlHAlignCenter
            
      .Columns("K").ColumnWidth = 15
      .Columns("L").ColumnWidth = 16
      .Columns("M").ColumnWidth = 15
      .Columns("N").ColumnWidth = 16
      .Columns("O").ColumnWidth = 18
                       
   End With
   
   g_rst_Princi.MoveFirst
     
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
   
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SOLMAE_CODINS))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
         
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "NUEVOS SOLES"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "DOLARES AMERICANOS"
      End If
            
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = 0
      End If
         
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = 0
      End If
      
      If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Or g_rst_Princi!SOLMAE_COMVTA_DOL > 0 Then
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_SOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00") * 100) + "%"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_DOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00") * 100) + "%"
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(0) + "%"
      End If
            
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = 0
      End If
         
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
      End If
              
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

Private Sub Limpia()
   ipp_FecIni.Text = (date)
   ipp_FecFin.Text = (date)
   Call gs_SetFocus(cmb_ConHip)
End Sub


