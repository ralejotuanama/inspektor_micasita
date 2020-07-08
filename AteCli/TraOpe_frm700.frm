VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RpIpk_01 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel3 
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   2990
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin EditLib.fpDateTime ipp_FecIni 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   240
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
         Left            =   1560
         TabIndex        =   9
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
      Begin VB.Label Label3 
         Caption         =   "Fecha Fin:"
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicio:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1065
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   1296
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   5040
         Picture         =   "TraOpe_frm700.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir de la Opción"
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   120
         Picture         =   "TraOpe_frm700.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exportar a Excel"
         Top             =   120
         Width           =   585
      End
      Begin Crystal.CrystalReport crp_Imprim 
         Left            =   840
         Top             =   240
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   1296
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel7 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Reporte de Inspektor"
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
         Left            =   120
         Picture         =   "TraOpe_frm700.frx":074C
         Top             =   120
         Width           =   480
      End
   End
End
Attribute VB_Name = "frm_RpIpk_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim l_str_Fecha      As String
Dim l_str_Hora       As String

Private Sub Limpia()
   ipp_FecIni.Text = (date - 30)
   ipp_FecFin.Text = (date)
End Sub

Private Sub fs_GenExc()

    Dim r_obj_Excel      As Excel.Application
    Dim r_int_ConVer     As Integer
    Dim r_dbl_GasAdm     As Double
    Dim r_dbl_GasFec     As Double
    Dim r_int_Cont       As Integer

    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " SELECT  C.INSPEK_FECCON, C.INPESK_USUCON, C.INSPEK_CODMOD, C.INSPEK_DOCCLI, C.INSPEK_RESULT,C.INSPEK_NOMBRE"
    g_str_Parame = g_str_Parame & "   FROM CRE_INSPEK C "
    g_str_Parame = g_str_Parame & "  WHERE "
    g_str_Parame = g_str_Parame & "C.INSPEK_FECCON >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
    g_str_Parame = g_str_Parame & "C.INSPEK_FECCON <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ""

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
   
       .Cells(2, 4) = "REPORTE INSPEKTOR " & CStr(ipp_FecIni.Text) & " Y EL " & CStr((ipp_FecFin.Text))
       .Cells(2, 2).Font.Bold = True
       .Cells(2, 2).HorizontalAlignment = xlHAlignCenter
      
       .Cells(4, 2) = "FECHA CONSULTA"
       .Cells(4, 3) = "USUARIO CONSULTA"
       .Cells(4, 4) = "COD. MODULO"
       .Cells(4, 5) = "DOC. CLIENTE"
       .Cells(4, 6) = "RESPUESTA"
       .Cells(4, 7) = "NOMBRES"
       
        .Range("B2:F2").Merge
        .Range("B2:F2").HorizontalAlignment = xlHAlignCenter
        .Range("B2:F2").Font.Bold = True
       
        .Range(.Cells(4, 1), .Cells(4, 15)).Font.Bold = True
        .Range(.Cells(4, 1), .Cells(4, 15)).HorizontalAlignment = xlHAlignCenter
           
        .Columns("A").HorizontalAlignment = xlHAlignCenter

        .Columns("B").ColumnWidth = 20
        .Columns("B").HorizontalAlignment = xlHAlignCenter
  
        .Columns("C").ColumnWidth = 20
        .Columns("C").HorizontalAlignment = xlHAlignCenter
     
        .Columns("D").ColumnWidth = 16
        .Columns("D").HorizontalAlignment = xlHAlignCenter
    
        .Columns("E").ColumnWidth = 16
        .Columns("E").HorizontalAlignment = xlHAlignCenter
      
        .Columns("F").ColumnWidth = 120
      
        .Columns("G").ColumnWidth = 35
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 5
   r_int_Cont = 1
   Do While Not g_rst_Princi.EOF
   
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = g_rst_Princi!INSPEK_FECCON
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = g_rst_Princi!INPESK_USUCON
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = g_rst_Princi!INSPEK_CODMOD
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = g_rst_Princi!INSPEK_DOCCLI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = g_rst_Princi!INSPEK_RESULT
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = g_rst_Princi!INSPEK_NOMBRE
   
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
   
Private Sub cmd_ExpExc_Click()
  'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   Call fs_GenExc
End Sub

Private Sub Form_Load()
 Me.Caption = modgen_g_str_NomPlt
 Call Limpia
   Call gs_CentraForm(Me)
End Sub

