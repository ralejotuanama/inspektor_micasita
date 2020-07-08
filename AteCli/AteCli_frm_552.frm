VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_RptSol_34 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   9465
   ClientTop       =   4125
   ClientWidth     =   5400
   Icon            =   "AteCli_frm_552.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2985
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5415
      _Version        =   65536
      _ExtentX        =   9551
      _ExtentY        =   5265
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
         TabIndex        =   6
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
            Height          =   225
            Left            =   660
            TabIndex        =   7
            Top             =   300
            Width           =   3855
            _Version        =   65536
            _ExtentX        =   6800
            _ExtentY        =   397
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   255
            Left            =   660
            TabIndex        =   12
            Top             =   30
            Width           =   3855
            _Version        =   65536
            _ExtentX        =   6800
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Cuadro de Seguimiento de Solicitudes"
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
            Picture         =   "AteCli_frm_552.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
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
            Left            =   30
            Picture         =   "AteCli_frm_552.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4740
            Picture         =   "AteCli_frm_552.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1515
         Left            =   30
         TabIndex        =   9
         Top             =   1440
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
         _ExtentY        =   2672
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
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   270
            Width           =   3225
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1500
            TabIndex        =   1
            Top             =   720
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
            Left            =   1500
            TabIndex        =   2
            Top             =   1080
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
         Begin VB.Label Label1 
            Caption         =   "Tipo de Reporte:"
            Height          =   255
            Left            =   150
            TabIndex        =   13
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   255
            Left            =   150
            TabIndex        =   11
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   225
            Left            =   150
            TabIndex        =   10
            Top             =   1110
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub cmb_TipRep_Click()
   If cmb_TipRep.ListIndex > -1 Then
      Select Case cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
         Case 1:
            ipp_FecIni.Text = Format("01/03/2007", "DD/MM/YYYY")
            ipp_FecFin.Text = (date)
            ipp_FecIni.Enabled = False
            ipp_FecFin.Enabled = True
         Case 2:
            ipp_FecIni.Text = (date)
            ipp_FecFin.Text = (date)
            ipp_FecIni.Enabled = True
            ipp_FecFin.Enabled = True
      End Select
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Call fs_GenExc
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Call Limpia
   Call gs_CentraForm(Me)
   Me.Caption = modgen_g_str_NomPlt
   Screen.MousePointer = 0
End Sub

Private Sub Limpia()
   ipp_FecIni.Text = (date)
   ipp_FecFin.Text = (date)
   ipp_FecIni.Enabled = False
   ipp_FecFin.Enabled = False
   
   cmb_TipRep.Clear
   cmb_TipRep.AddItem "A UNA FECHA"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
   cmb_TipRep.AddItem "POR RANGO DE FECHAS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2
   cmb_TipRep.ListIndex = -1
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

Private Sub fs_GenExc()
   Dim r_int_InsDia001     As Integer
   Dim r_int_InsDia002     As Integer
   Dim r_int_InsDia003     As Integer
   Dim r_int_InsDia004     As Integer
   Dim r_int_InsDia005     As Integer
   Dim r_int_InsDia006     As Integer
   Dim r_int_InsDia007     As Integer
   Dim r_int_InsMes001     As Integer
   Dim r_int_InsMes002     As Integer
   Dim r_int_InsMes003     As Integer
   Dim r_int_InsMes004     As Integer
   Dim r_int_InsMes005     As Integer
   Dim r_int_InsMes006     As Integer
   Dim r_int_InsMes007     As Integer
   Dim r_int_InsAcu001     As Integer
   Dim r_int_InsAcu002     As Integer
   Dim r_int_InsAcu003     As Integer
   Dim r_int_InsAcu004     As Integer
   Dim r_int_InsAcu005     As Integer
   Dim r_int_InsAcu006     As Integer
   Dim r_int_InsAcu007     As Integer
   Dim r_str_Titulo        As String
   Dim r_obj_Excel         As Excel.Application
   
   'Configura el excel
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   r_str_Titulo = "SEGUIMIENTO DE SOLICITUDES AL " & Format(date, "dd/mm/yyyy")
   
   With r_obj_Excel.ActiveSheet
      .Cells(2, 3) = r_str_Titulo
      
      .Cells(4, 2) = "INSTANCIA DE EVALUACION"
      .Cells(6, 2) = "ATENCION COMERCIAL"
      .Cells(7, 2) = "EVAL./APROB. CREDITICIA Y DOC. CLIENTE"
      .Cells(8, 2) = "PAGO DE GASTOS CIERRE"
      .Cells(9, 2) = "TASACION Y SEGUROS"
      .Cells(10, 2) = "LEGAL"
      .Cells(11, 2) = "POLIZAS Y COFIDE"
      .Cells(12, 2) = "RECHAZOS AUTOMATICOS"
      .Cells(14, 2) = "TOTALES"
      
      .Cells(4, 3) = "MOVIMIENTOS"
      .Cells(5, 3) = "DIARIO"
      .Cells(5, 4) = "MENSUAL"
      .Cells(5, 5) = "ACUMULADO"
      
      .Range(.Cells(2, 1), .Cells(2, 5)).Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 3), .Cells(4, 5)).Font.Bold = True
      .Range(.Cells(4, 3), .Cells(4, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 2), .Cells(5, 2)).Font.Bold = True
      .Range(.Cells(4, 2), .Cells(5, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 2), .Cells(5, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(5, 3), .Cells(5, 5)).Font.Bold = True
      .Range(.Cells(5, 3), .Cells(5, 5)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(6, 3), .Cells(14, 5)).HorizontalAlignment = xlHAlignCenter
      
      'Une las celdas
      .Range("B2:E2").Merge
      .Range("C4:E4").Merge
      .Range("B4:B5").Merge
            
      'Ancho de columnas
      .Columns("B").ColumnWidth = 40
      .Columns("C").ColumnWidth = 14
      .Columns("D").ColumnWidth = 14
      .Columns("E").ColumnWidth = 14
   End With
   
   'Obtiene el total diario de Atencion Comercial (001)
   Call ff_ObtieneTotales_001(r_int_InsDia001, r_int_InsMes001, r_int_InsAcu001)
   
   'Obtiene el total diario de Evaluacion y Aprobacion Crediticia (002)
   Call ff_ObtieneTotales_002(r_int_InsDia002, r_int_InsMes002, r_int_InsAcu002)
   
   'Obtiene el total diario de Pago de Gastos de Cierre (003)
   Call ff_ObtieneTotales_003(r_int_InsDia003, r_int_InsMes003, r_int_InsAcu003)
   
   'Obtiene el total diario de Tasacion y Seguros (004)
   Call ff_ObtieneTotales_004(r_int_InsDia004, r_int_InsMes004, r_int_InsAcu004)
   
   'Obtiene el total diario de Legal (005)
   Call ff_ObtieneTotales_005(r_int_InsDia005, r_int_InsMes005, r_int_InsAcu005)
   
   'Obtiene el total diario de Polizas y Cofide (006)
   Call ff_ObtieneTotales_006(r_int_InsDia006, r_int_InsMes006, r_int_InsAcu006)
   
   'Obtiene los rechazos automaticos (007)
   Call ff_ObtieneTotales_007(r_int_InsDia007, r_int_InsMes007, r_int_InsAcu007)
   
   r_obj_Excel.ActiveSheet.Cells(6, 3) = Format(r_int_InsDia001, "000")
   r_obj_Excel.ActiveSheet.Cells(7, 3) = Format(r_int_InsDia002, "000")
   r_obj_Excel.ActiveSheet.Cells(8, 3) = Format(r_int_InsDia003, "000")
   r_obj_Excel.ActiveSheet.Cells(9, 3) = Format(r_int_InsDia004, "000")
   r_obj_Excel.ActiveSheet.Cells(10, 3) = Format(r_int_InsDia005, "000")
   r_obj_Excel.ActiveSheet.Cells(11, 3) = Format(r_int_InsDia006, "000")
   r_obj_Excel.ActiveSheet.Cells(12, 3) = Format(r_int_InsDia007, "000")
   r_obj_Excel.ActiveSheet.Cells(6, 4) = Format(r_int_InsMes001, "000")
   r_obj_Excel.ActiveSheet.Cells(7, 4) = Format(r_int_InsMes002, "000")
   r_obj_Excel.ActiveSheet.Cells(8, 4) = Format(r_int_InsMes003, "000")
   r_obj_Excel.ActiveSheet.Cells(9, 4) = Format(r_int_InsMes004, "000")
   r_obj_Excel.ActiveSheet.Cells(10, 4) = Format(r_int_InsMes005, "000")
   r_obj_Excel.ActiveSheet.Cells(11, 4) = Format(r_int_InsMes006, "000")
   r_obj_Excel.ActiveSheet.Cells(12, 4) = Format(r_int_InsMes007, "000")
   r_obj_Excel.ActiveSheet.Cells(6, 5) = Format(r_int_InsAcu001, "000")
   r_obj_Excel.ActiveSheet.Cells(7, 5) = Format(r_int_InsAcu002, "000")
   r_obj_Excel.ActiveSheet.Cells(8, 5) = Format(r_int_InsAcu003, "000")
   r_obj_Excel.ActiveSheet.Cells(9, 5) = Format(r_int_InsAcu004, "000")
   r_obj_Excel.ActiveSheet.Cells(10, 5) = Format(r_int_InsAcu005, "000")
   r_obj_Excel.ActiveSheet.Cells(11, 5) = Format(r_int_InsAcu006, "000")
   r_obj_Excel.ActiveSheet.Cells(12, 5) = Format(r_int_InsAcu007, "000")
   r_obj_Excel.ActiveSheet.Cells(14, 3) = Format(r_int_InsDia001 + r_int_InsDia002 + r_int_InsDia003 + r_int_InsDia004 + r_int_InsDia005 + r_int_InsDia006 + r_int_InsDia007, "000")
   r_obj_Excel.ActiveSheet.Cells(14, 4) = Format(r_int_InsMes001 + r_int_InsMes002 + r_int_InsMes003 + r_int_InsMes004 + r_int_InsMes005 + r_int_InsMes006 + r_int_InsMes007, "000")
   r_obj_Excel.ActiveSheet.Cells(14, 5) = Format(r_int_InsAcu001 + r_int_InsAcu002 + r_int_InsAcu003 + r_int_InsAcu004 + r_int_InsAcu005 + r_int_InsAcu006 + r_int_InsAcu007, "000")
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub ff_ObtieneTotales_001(ByRef p_InsDia001 As Integer, ByRef p_InsMes001 As Integer, ByRef p_InsAcu001 As Integer)
   p_InsDia001 = 0
   p_InsMes001 = 0
   p_InsAcu001 = 0
   
   g_str_Parame = "SELECT SOLMAE_FECSOL FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = 11 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1  AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         'Acumulado diario
         If Trim(CStr(g_rst_Listas!SOLMAE_FECSOL)) = Trim(CStr(Format(date, "yyyymmdd"))) Then
            p_InsDia001 = p_InsDia001 + 1
         End If
         
         'Acumulado por el mes actual
         If Trim(CStr(Mid(g_rst_Listas!SOLMAE_FECSOL, 1, 6))) = Trim(CStr(Format(date, "yyyymm"))) Then
            p_InsMes001 = p_InsMes001 + 1
         End If
         
         'Acumulado total por el rango seleccionado
         p_InsAcu001 = p_InsAcu001 + 1
         
         'Siguiente registros
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub ff_ObtieneTotales_002(ByRef p_InsDia002 As Integer, ByRef p_InsMes002 As Integer, ByRef p_InsAcu002 As Integer)
   p_InsDia002 = 0
   p_InsMes002 = 0
   p_InsAcu002 = 0
   
   g_str_Parame = "SELECT SOLMAE_FECSOL FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "(SOLMAE_CODINS = 21 OR SOLMAE_CODINS = 31 OR SOLMAE_CODINS = 32) AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         'Acumulado diario
         If Trim(CStr(g_rst_Listas!SOLMAE_FECSOL)) = Trim(CStr(Format(date, "yyyymmdd"))) Then
            p_InsDia002 = p_InsDia002 + 1
         End If
         
         'Acumulado por el mes actual
         If Trim(CStr(Mid(g_rst_Listas!SOLMAE_FECSOL, 1, 6))) = Trim(CStr(Format(date, "yyyymm"))) Then
            p_InsMes002 = p_InsMes002 + 1
         End If
         
         'Acumulado total por el rango seleccionado
         p_InsAcu002 = p_InsAcu002 + 1
         
         'Siguiente registros
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub ff_ObtieneTotales_003(ByRef p_InsDia003 As Integer, ByRef p_InsMes003 As Integer, ByRef p_InsAcu003 As Integer)
   p_InsDia003 = 0
   p_InsMes003 = 0
   p_InsAcu003 = 0
   
   g_str_Parame = "SELECT SOLMAE_NUMERO, SOLMAE_FECSOL FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS > 21 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1  AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         
         'Valida si pagos los gasto de cierre
         If ff_GasAdm(g_rst_Listas!SOLMAE_NUMERO) > 0 Then
         
            'Acumulado diario
            If Trim(CStr(g_rst_Listas!SOLMAE_FECSOL)) = Trim(CStr(Format(date, "yyyymmdd"))) Then
               p_InsDia003 = p_InsDia003 + 1
            End If
            
            'Acumulado por el mes actual
            If Trim(CStr(Mid(g_rst_Listas!SOLMAE_FECSOL, 1, 6))) = Trim(CStr(Format(date, "yyyymm"))) Then
               p_InsMes003 = p_InsMes003 + 1
            End If
            
            'Acumulado total por el rango seleccionado
            p_InsAcu003 = p_InsAcu003 + 1
            
         End If
         
         'Siguiente registros
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub ff_ObtieneTotales_004(ByRef p_InsDia004 As Integer, ByRef p_InsMes004 As Integer, ByRef p_InsAcu004 As Integer)
   p_InsDia004 = 0
   p_InsMes004 = 0
   p_InsAcu004 = 0
   
   g_str_Parame = "SELECT SOLMAE_FECSOL FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "(SOLMAE_CODINS = 41 OR SOLMAE_CODINS = 42) AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         'Acumulado diario
         If Trim(CStr(g_rst_Listas!SOLMAE_FECSOL)) = Trim(CStr(Format(date, "yyyymmdd"))) Then
            p_InsDia004 = p_InsDia004 + 1
         End If
         
         'Acumulado por el mes actual
         If Trim(CStr(Mid(g_rst_Listas!SOLMAE_FECSOL, 1, 6))) = Trim(CStr(Format(date, "yyyymm"))) Then
            p_InsMes004 = p_InsMes004 + 1
         End If
         
         'Acumulado total por el rango seleccionado
         p_InsAcu004 = p_InsAcu004 + 1
         
         'Siguiente registros
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub ff_ObtieneTotales_005(ByRef p_InsDia005 As Integer, ByRef p_InsMes005 As Integer, ByRef p_InsAcu005 As Integer)
   p_InsDia005 = 0
   p_InsMes005 = 0
   p_InsAcu005 = 0
   
   g_str_Parame = "SELECT SOLMAE_FECSOL FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = 51 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         'Acumulado diario
         If Trim(CStr(g_rst_Listas!SOLMAE_FECSOL)) = Trim(CStr(Format(date, "yyyymmdd"))) Then
            p_InsDia005 = p_InsDia005 + 1
         End If
         
         'Acumulado por el mes actual
         If Trim(CStr(Mid(g_rst_Listas!SOLMAE_FECSOL, 1, 6))) = Trim(CStr(Format(date, "yyyymm"))) Then
            p_InsMes005 = p_InsMes005 + 1
         End If
         
         'Acumulado total por el rango seleccionado
         p_InsAcu005 = p_InsAcu005 + 1
         
         'Siguiente registros
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub ff_ObtieneTotales_006(ByRef p_InsDia006 As Integer, ByRef p_InsMes006 As Integer, ByRef p_InsAcu006 As Integer)
   p_InsDia006 = 0
   p_InsMes006 = 0
   p_InsAcu006 = 0
   
   g_str_Parame = "SELECT SOLMAE_FECSOL FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "(SOLMAE_CODINS = 61 OR SOLMAE_CODINS = 62) AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         'Acumulado diario
         If Trim(CStr(g_rst_Listas!SOLMAE_FECSOL)) = Trim(CStr(Format(date, "yyyymmdd"))) Then
            p_InsDia006 = p_InsDia006 + 1
         End If
         
         'Acumulado por el mes actual
         If Trim(CStr(Mid(g_rst_Listas!SOLMAE_FECSOL, 1, 6))) = Trim(CStr(Format(date, "yyyymm"))) Then
            p_InsMes006 = p_InsMes006 + 1
         End If
         
         'Acumulado total por el rango seleccionado
         p_InsAcu006 = p_InsAcu006 + 1
         
         'Siguiente registros
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub ff_ObtieneTotales_007(ByRef p_InsDia007 As Integer, ByRef p_InsMes007 As Integer, ByRef p_InsAcu007 As Integer)
   p_InsDia007 = 0
   p_InsMes007 = 0
   p_InsAcu007 = 0
   
   g_str_Parame = "SELECT SEGFECCRE FROM TRA_RECADM WHERE "
   g_str_Parame = g_str_Parame & "SEGPLTCRE = 'PRDP0001' AND "
   g_str_Parame = g_str_Parame & "SEGFECCRE >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SEGFECCRE <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         'Acumulado diario
         If Trim(CStr(g_rst_Listas!SEGFECCRE)) = Trim(CStr(Format(date, "yyyymmdd"))) Then
            p_InsDia007 = p_InsDia007 + 1
         End If
         
         'Acumulado por el mes actual
         If Trim(CStr(Mid(g_rst_Listas!SEGFECCRE, 1, 6))) = Trim(CStr(Format(date, "yyyymm"))) Then
            p_InsMes007 = p_InsMes007 + 1
         End If
         
         'Acumulado total por el rango seleccionado
         p_InsAcu007 = p_InsAcu007 + 1
         
         'Siguiente registros
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Function ff_GasAdm(ByVal p_NumSol As String, Optional ByRef p_FecPag As Double) As Double
Dim r_rst_PagGto  As Recordset
   
   ff_GasAdm = 0
   p_FecPag = 0
   
   g_str_Parame = "SELECT GASADM_PAGIMP, GASADM_PAGFEC FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "GASADM_SITUAC = 1"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_PagGto, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_PagGto.BOF And r_rst_PagGto.EOF) Then
      r_rst_PagGto.MoveFirst
      
      Do While Not r_rst_PagGto.EOF
         ff_GasAdm = ff_GasAdm + r_rst_PagGto!GASADM_PAGIMP
         p_FecPag = r_rst_PagGto!GASADM_PAGFEC
         r_rst_PagGto.MoveNext
      Loop
   End If
   
   r_rst_PagGto.Close
   Set r_rst_PagGto = Nothing
End Function

