VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_RptSol_36 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "AteCli_frm_560.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   5442
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   1
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
            TabIndex        =   2
            Top             =   30
            Width           =   3405
            _Version        =   65536
            _ExtentX        =   6006
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Reporte de Prospectos"
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
            TabIndex        =   3
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
            Picture         =   "AteCli_frm_560.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   4
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
            Picture         =   "AteCli_frm_560.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4740
            Picture         =   "AteCli_frm_560.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1515
         Left            =   30
         TabIndex        =   7
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
         Begin VB.CheckBox chk_ConHip 
            Caption         =   "Todos los Consejero Hipotecario"
            Height          =   315
            Left            =   1230
            TabIndex        =   9
            Top             =   420
            Width           =   2685
         End
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            ItemData        =   "AteCli_frm_560.frx":0A62
            Left            =   1230
            List            =   "AteCli_frm_560.frx":0A64
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   60
            Width           =   4065
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1230
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
            Left            =   1230
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
         Begin VB.Label Label4 
            Caption         =   "Consejero Hipotecario:"
            Height          =   465
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1005
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   255
            Left            =   60
            TabIndex        =   13
            Top             =   780
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   225
            Left            =   60
            TabIndex        =   12
            Top             =   1110
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_ConHip()      As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub cmd_ExpExc_Click()
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
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
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
   
   Call Limpia
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_ConHip)
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_dbl_GasAdm     As Double
Dim r_dbl_GasFec     As Double
Dim r_int_Cont       As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT PROMAE_NUMDOC , PROMAE_TIPDOC , PROMAE_NUMDOC, "
   g_str_Parame = g_str_Parame & "       PROCLI_APEPAT , PROCLI_APEMAT , PROCLI_NOMBRE, "
   g_str_Parame = g_str_Parame & "       SOLMAE_TITTDO , SOLMAE_TITNDO , PROMAE_CODCON, PROMAE_PROYEC, PROMAE_PROMOT, PROMAE_CONSTR, POSMAE_FECCON,"
   g_str_Parame = g_str_Parame & "       PROMAE_FECCON , SOLMAE_FECSOL , SOLMAE_SITUAC, "
   g_str_Parame = g_str_Parame & "       PARDES_DESCRI , SOLMAE_CONHIP , SOLMAE_CODPRD  "
   g_str_Parame = g_str_Parame & "  FROM CRE_PROMAE A "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLMAE B ON TRIM(B.SOLMAE_TITTDO) = TRIM(A.PROMAE_TIPDOC) AND TRIM(B.SOLMAE_TITNDO) = TRIM(A.PROMAE_NUMDOC) "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_PROCLI C ON C.PROCLI_TIPDOC = A.PROMAE_TIPDOC AND C.PROCLI_NUMDOC = A.PROMAE_NUMDOC "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES D ON D.PARDES_CODITE = B.SOLMAE_SITUAC AND D.PARDES_CODGRP = '020' AND PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_POSMAE E ON E.POSMAE_TIPDOC = C.PROCLI_TIPDOC AND E.POSMAE_NUMDOC = C.PROCLI_NUMDOC "
   g_str_Parame = g_str_Parame & "  WHERE A.PROMAE_FECCON BETWEEN " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   If chk_ConHip.Value = 0 Then
      g_str_Parame = g_str_Parame & " AND PROMAE_CODCON = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "' "
   End If
   g_str_Parame = g_str_Parame & " ORDER BY PROMAE_FECCON "
   
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
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      .Range(.Cells(2, 1), .Cells(2, 12)).Merge
      .Range(.Cells(2, 1), .Cells(2, 12)).Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 12)).HorizontalAlignment = xlHAlignCenter
      .Cells(2, 1) = "REPORTE DE CLIENTES POTENCIALES Y PROSPECTOS REGISTRADOS ENTRE EL " & CStr(ipp_FecIni.Text) & " Y EL " & CStr((ipp_FecFin.Text))
      
      .Cells(4, 1) = "ITEM"
      .Cells(4, 2) = "FECHA CLIENTE POT."
      .Cells(4, 3) = "FECHA PROSPECTO"
      .Cells(4, 4) = "DOC. IDENTIDAD"
      .Cells(4, 5) = "NOMBRE PROSPECTO"
      .Cells(4, 6) = "PROYECTO"
      .Cells(4, 7) = "PROMOTOR"
      .Cells(4, 8) = "CONSTRUCTOR"
      .Cells(4, 9) = "CONSEJERO PROS."
      .Cells(4, 10) = "SITUACION"
      .Cells(4, 11) = "CONSEJERO FINAL "
      .Cells(4, 12) = "PRODUCTO"
      
      .Range(.Cells(4, 1), .Cells(4, 15)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 15)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 5
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      
      .Columns("B").ColumnWidth = 17
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 17
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 17
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 40
      
      .Columns("F").ColumnWidth = 45
      
      .Columns("G").ColumnWidth = 35
            
      .Columns("H").ColumnWidth = 35
      
      .Columns("I").ColumnWidth = 15
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 22
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 22
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 40
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      
      .Range("A1:L1000000").Font.Name = "Arial"
      .Range("A1:L1000000").Font.Size = 8
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 5
   r_int_Cont = 1
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont
            
      If IsNull(g_rst_Princi!POSMAE_FECCON) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = ""
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "'" & Right(CStr(g_rst_Princi!POSMAE_FECCON), 2) & "/" & Mid(CStr(g_rst_Princi!POSMAE_FECCON), 5, 2) & "/" & Left(CStr(g_rst_Princi!POSMAE_FECCON), 4)
      End If
      
      If IsNull(g_rst_Princi!PROMAE_FECCON) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = ""
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = "'" & Right(CStr(g_rst_Princi!PROMAE_FECCON), 2) & "/" & Mid(CStr(g_rst_Princi!PROMAE_FECCON), 5, 2) & "/" & Left(CStr(g_rst_Princi!PROMAE_FECCON), 4)
      End If
            
'      If IsNull(g_rst_Princi!SOLMAE_FECSOL) Then
'         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = ""
'      Else
'         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = "'" & Right(CStr(g_rst_Princi!SOLMAE_FECSOL), 2) & "/" & Mid(CStr(g_rst_Princi!SOLMAE_FECSOL), 5, 2) & "/" & Left(CStr(g_rst_Princi!SOLMAE_FECSOL), 4)
'      End If

      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = g_rst_Princi!PROMAE_NUMDOC
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!PROCLI_APEPAT) & " " & Trim(g_rst_Princi!PROCLI_APEMAT) & " " & Trim(g_rst_Princi!PROCLI_NOMBRE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = g_rst_Princi!PROMAE_PROYEC
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = g_rst_Princi!PROMAE_PROMOT
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = g_rst_Princi!PROMAE_CONSTR
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = "" & Trim(g_rst_Princi!PROMAE_CODCON)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "" & Trim(g_rst_Princi!PARDES_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = "" & Trim(g_rst_Princi!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = moddat_gf_Consulta_Produc("" & g_rst_Princi!SOLMAE_CODPRD)
      
                    
      r_int_ConVer = r_int_ConVer + 1
      r_int_Cont = r_int_Cont + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
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
   ipp_FecIni.Text = (date - 30)
   ipp_FecFin.Text = (date)
   Call gs_SetFocus(cmb_ConHip)
End Sub

