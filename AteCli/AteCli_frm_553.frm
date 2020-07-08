VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_RptSol_35 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   5085
   ClientTop       =   4260
   ClientWidth     =   5370
   Icon            =   "AteCli_frm_553.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel6 
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
         Height          =   285
         Left            =   660
         TabIndex        =   1
         Top             =   30
         Width           =   4485
         _Version        =   65536
         _ExtentX        =   7911
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Reporte de Solicitudes por Fechas "
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   255
         Left            =   660
         TabIndex        =   2
         Top             =   300
         Width           =   3345
         _Version        =   65536
         _ExtentX        =   5900
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Por tiempo de proceso en instancias"
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
         Picture         =   "AteCli_frm_553.frx":000C
         Top             =   60
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   720
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
         Left            =   60
         Picture         =   "AteCli_frm_553.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exportar a Excel"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   4740
         Picture         =   "AteCli_frm_553.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   1755
      Left            =   0
      TabIndex        =   6
      Top             =   1410
      Width           =   5355
      _Version        =   65536
      _ExtentX        =   9446
      _ExtentY        =   3096
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
      Begin VB.CheckBox Chk_FecPro 
         Caption         =   "Todas las Solicitudes en Trámite"
         Height          =   315
         Left            =   1020
         TabIndex        =   9
         Top             =   1410
         Width           =   4245
      End
      Begin VB.CheckBox chk_Produc 
         Caption         =   "Todos los Productos"
         Height          =   315
         Left            =   1020
         TabIndex        =   8
         Top             =   420
         Width           =   2685
      End
      Begin VB.ComboBox cmb_TipPro 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   60
         Width           =   4275
      End
      Begin EditLib.fpDateTime ipp_FecIni 
         Height          =   315
         Left            =   1020
         TabIndex        =   10
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
         Left            =   1020
         TabIndex        =   11
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
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicio:"
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Fin:"
         Height          =   225
         Left            =   60
         TabIndex        =   13
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   765
      End
   End
End
Attribute VB_Name = "frm_RptSol_35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub Chk_FecPro_Click()
   If Chk_FecPro.Value = 1 Then
      ipp_FecIni.Enabled = False
      ipp_FecFin.Enabled = False
      Call gs_SetFocus(cmd_ExpExc)
   ElseIf Chk_FecPro.Value = 0 Then
      ipp_FecIni.Enabled = True
      ipp_FecFin.Enabled = True
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub chk_Produc_Click()
   If chk_Produc.Value = 1 Then
      cmb_TipPro.ListIndex = -1
      cmb_TipPro.Enabled = False
      Call gs_SetFocus(ipp_FecIni)
   ElseIf chk_Produc.Value = 0 Then
      cmb_TipPro.Enabled = True
      Call gs_SetFocus(cmb_TipPro)
   End If
End Sub

Private Sub cmb_TipPro_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_TipPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPro_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   'Validación
   If chk_Produc.Value = 0 Then
      If cmb_TipPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPro)
         Exit Sub
      End If
   End If
   If cmb_TipPro.ListIndex <> -1 Then
      If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
         MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIni)
         Exit Sub
      End If
   End If
   If Chk_FecPro.Value = 0 Then
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
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   ipp_FecIni.Text = date - 30
   ipp_FecFin.Text = date
   Call moddat_gs_Carga_Produc(cmb_TipPro, l_arr_Produc, 4)
   Call gs_CentraForm(Me)
   Call gs_SetFocus(cmb_TipPro)
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_bol_Muestra    As Boolean
    
   g_str_Parame = "USP_CUR_SEGINST ("
   If Chk_FecPro.Value = 0 Then
       g_str_Parame = g_str_Parame & "'" & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "', "
       g_str_Parame = g_str_Parame & "'" & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & "', "
   Else
       g_str_Parame = g_str_Parame & "'0', '0', "
   End If
   If chk_Produc.Value = 0 Then
       g_str_Parame = g_str_Parame & "'" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "') "
   Else
       g_str_Parame = g_str_Parame & "'') "
   End If
      
   'EJECUTA CONSULTA
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CUR_SEGINST.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
       
   'Si no encuentra ninguna Solicitud
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_ConVer = 2
    
   With r_obj_Excel.ActiveSheet
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 15)).Merge
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 15)).Font.Bold = True
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 15)).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_ConVer, 1) = "TIEMPO DE PROCESO POR INSTANCIA DEL " & CDate(ipp_FecIni.Text) & " AL " & CDate(ipp_FecFin.Text)
      r_int_ConVer = r_int_ConVer + 3
      
      .Cells(r_int_ConVer, 1) = "ITEM"
      .Cells(r_int_ConVer, 2) = "SOLICITUD"
      .Cells(r_int_ConVer, 3) = "NOMBRE CLIENTE"
      .Cells(r_int_ConVer, 4) = "DOC. IDENTIDAD"
      .Cells(r_int_ConVer, 5) = "F. SOLICITUD"
      .Cells(r_int_ConVer, 6) = "ATEN. COMERCIAL"
      .Cells(r_int_ConVer, 7) = "EVAL. CREDITICIA"
      .Cells(r_int_ConVer, 8) = "APROB. CREDITICIA"
      .Cells(r_int_ConVer, 9) = "DOC. DEL INMUEBLE"
      .Cells(r_int_ConVer, 10) = "TASACIÓN"
      .Cells(r_int_ConVer, 11) = "EVAL. SEGUROS"
      .Cells(r_int_ConVer, 12) = "EVAL. LEGAL"
      .Cells(r_int_ConVer, 13) = "TRÁMITE POLIZAS"
      .Cells(r_int_ConVer, 14) = "TRÁMITE COFIDE"
      .Cells(r_int_ConVer, 15) = "SITUACIÓN"
     
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 15)).Font.Bold = True
      
      .Columns("A").ColumnWidth = 5
      
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 36
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 13
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 17
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 16
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 18
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 18
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 11
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 15
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 12
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      
      .Columns("M").ColumnWidth = 16
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      
      .Columns("N").ColumnWidth = 16
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      
      .Columns("O").ColumnWidth = 16
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 15)).HorizontalAlignment = xlHAlignCenter
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = r_int_ConVer + 1
    
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 5
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = g_rst_Princi!NROSOL
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = g_rst_Princi!CLIENTE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!NRODOC)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CDate(g_rst_Princi!FECSOL)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CStr(g_rst_Princi!INSTAN01) & "  (" & CStr(g_rst_Princi!INSOBS01) & ")"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CStr(g_rst_Princi!INSTAN02) & "  (" & CStr(g_rst_Princi!INSOBS02) & ")"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CStr(g_rst_Princi!INSTAN03) & "  (" & CStr(g_rst_Princi!INSOBS03) & ")"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CStr(g_rst_Princi!INSTAN04) & "  (" & CStr(g_rst_Princi!INSOBS04) & ")"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = CStr(g_rst_Princi!INSTAN05) & "  (" & CStr(g_rst_Princi!INSOBS05) & ")"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = CStr(g_rst_Princi!INSTAN06) & "  (" & CStr(g_rst_Princi!INSOBS06) & ")"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CStr(g_rst_Princi!INSTAN07) & "  (" & CStr(g_rst_Princi!INSOBS07) & ")"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(g_rst_Princi!INSTAN08) & "  (" & CStr(g_rst_Princi!INSOBS08) & ")"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = CStr(g_rst_Princi!INSTAN09) & "  (" & CStr(g_rst_Princi!INSOBS09) & ")"
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = CStr(g_rst_Princi!STUACION)
      
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
      Call gs_SetFocus(cmd_ExpExc)
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
   Call gs_SetFocus(cmb_TipPro)
End Sub


