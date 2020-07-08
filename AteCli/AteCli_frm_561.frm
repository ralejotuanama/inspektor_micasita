VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_RptSol_37 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   Icon            =   "AteCli_frm_561.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2790
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   4921
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
         Left            =   60
         TabIndex        =   6
         Top             =   60
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
            Height          =   525
            Left            =   660
            TabIndex        =   7
            Top             =   30
            Width           =   4245
            _Version        =   65536
            _ExtentX        =   7488
            _ExtentY        =   926
            _StockProps     =   15
            Caption         =   "Reporte de Desembolsos"
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
            Picture         =   "AteCli_frm_561.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   8
         Top             =   780
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
            Picture         =   "AteCli_frm_561.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4740
            Picture         =   "AteCli_frm_561.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1260
         Left            =   60
         TabIndex        =   9
         Top             =   1470
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
         _ExtentY        =   2222
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
            ItemData        =   "AteCli_frm_561.frx":0A62
            Left            =   1305
            List            =   "AteCli_frm_561.frx":0A6C
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   135
            Width           =   3120
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   330
            Left            =   1305
            TabIndex        =   1
            Top             =   495
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
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
            BorderColor     =   0
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
            Height          =   330
            Left            =   1305
            TabIndex        =   2
            Top             =   855
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
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
         Begin VB.Label Label1 
            Caption         =   "Fecha Fin:"
            Height          =   240
            Left            =   135
            TabIndex        =   12
            Top             =   855
            Width           =   1050
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Reporte:"
            Height          =   240
            Left            =   135
            TabIndex        =   11
            Top             =   135
            Width           =   1050
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   285
            Left            =   135
            TabIndex        =   10
            Top             =   495
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ExpExc_Click()
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   If cmb_TipRep.ListIndex = 0 Then
      Call fs_GenExc_Promotora
   Else
      Call fs_GenExc_Consejero
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(ipp_FecIni)
   Screen.MousePointer = 0
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIni)
   End If
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
   cmb_TipRep.ListIndex = 0
End Sub

Private Sub fs_GenExc_Promotora()
Dim r_rst_Bucle      As ADODB.Recordset
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_Cont       As Integer
Dim r_int_Cont1      As Integer
Dim r_int_ConVer1    As Integer
Dim r_str_Nombre     As String
Dim r_int_ProdMV     As Integer
Dim r_int_ProdMC     As Integer
Dim r_int_ProdCC     As Integer
Dim r_int_TotProd    As Integer
Dim r_dbl_MontoMV    As Double
Dim r_dbl_MontoMC    As Double
Dim r_dbl_MontoCC    As Double
Dim r_dbl_MontoTot   As Double
Dim r_int_TotalMV    As Integer
Dim r_int_TotalMC    As Integer
Dim r_int_TotalCC    As Integer
Dim r_dbl_MonTotMV   As Double
Dim r_dbl_MonTotMC   As Double
Dim r_dbl_MonTotCC   As Double
Dim r_int_Porcent    As Integer
Dim r_int_Cantidad   As Integer
Dim r_int_Total      As Integer
Dim r_dbl_TotGrpPor  As Double
Dim r_dbl_TotGrp     As Double

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "DETALLE DE DESEMBOLSOS POR PROMOTORAS DEL " & CStr(ipp_FecIni.Text) & " AL " & CStr((ipp_FecFin.Text))
      .Range("A1:L1").Select
      .Range("A1:L1").HorizontalAlignment = xlHAlignCenter
      .Range("A1:L1").Font.Bold = True
      r_obj_Excel.Selection.MergeCells = True

      .Cells(3, 1) = "ITEM"
      .Cells(3, 2) = "PROMOTORAS"
      .Cells(3, 3) = "PROYECTOS"
      .Cells(4, 4) = "MiVivienda"
      .Cells(4, 5) = "MiCasita"
      .Cells(4, 6) = "Coficasa"
      .Cells(3, 6) = "PRODUCTO"
      .Cells(4, 7) = "TOTAL POR PRODUCTO"
      .Cells(4, 8) = "%"
      .Cells(4, 9) = "MiVivienda"
      .Cells(4, 10) = "MiCasita"
      .Cells(4, 11) = "Coficasa"
      .Cells(3, 11) = "MONTO"
      .Cells(4, 12) = "TOTAL POR MONTO"

      .Range("A3:A4").Select
      r_obj_Excel.Selection.MergeCells = True
      .Range("B3:B4").Select
      r_obj_Excel.Selection.MergeCells = True
      .Range("C3:C4").Select
      r_obj_Excel.Selection.MergeCells = True
      .Range("D3:F3").Select
      r_obj_Excel.Selection.MergeCells = True
      
      .Range("G3:G4").Select
      r_obj_Excel.Selection.MergeCells = True
      r_obj_Excel.Selection.Cells.WrapText = True
      .Range("H3:H4").Select
      r_obj_Excel.Selection.MergeCells = True
      r_obj_Excel.Selection.Cells.WrapText = True
      .Range("I3:K3").Select
      r_obj_Excel.Selection.MergeCells = True
      .Range("L3:L4").Select
      r_obj_Excel.Selection.MergeCells = True
      r_obj_Excel.Selection.Cells.WrapText = True

      .Range(.Cells(3, 1), .Cells(3, 12)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(3, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 12)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 12)).HorizontalAlignment = xlHAlignCenter

      .Columns("A").ColumnWidth = 5
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 50
      .Columns("C").ColumnWidth = 50
      .Columns("D").ColumnWidth = 11
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 11
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 11
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 12
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 10
      .Columns("I").ColumnWidth = 12
      .Columns("J").ColumnWidth = 12
      .Columns("K").ColumnWidth = 12
      .Columns("L").ColumnWidth = 14
   End With

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT COUNT(A.HIPMAE_CODPRD) AS CONTGRUPO "
   g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & "   LEFT JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.HIPMAE_PRYINM "
   g_str_Parame = g_str_Parame & "   LEFT JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = B.DATGEN_VENTDO AND C.DATGEN_EMPNDO = B.DATGEN_VENNDO "
   g_str_Parame = g_str_Parame & "  WHERE A.HIPMAE_SITUAC IN (2,6,9) AND A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "  AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_Porcent = g_rst_Princi!CONTGRUPO
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Para ser tomado por consejero en el proximo query
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT CASE WHEN C.DATGEN_RAZSOC IS NULL THEN 'RECURSOS PROPIOS' ELSE TRIM(C.DATGEN_RAZSOC) END AS PROMOTORA, "
   g_str_Parame = g_str_Parame & "       COUNT(CASE WHEN A.HIPMAE_CODPRD <> '011' THEN A.HIPMAE_CODPRD END) + COUNT(CASE WHEN A.HIPMAE_CODPRD =  '011' THEN A.HIPMAE_CODPRD END) TOTAL_PRODUCTO "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.HIPMAE_PRYINM "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = B.DATGEN_VENTDO AND C.DATGEN_EMPNDO = B.DATGEN_VENNDO "
   g_str_Parame = g_str_Parame & " WHERE A.HIPMAE_SITUAC IN (2,6,9) AND A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "  AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & " GROUP BY C.DATGEN_RAZSOC ORDER BY 2 DESC"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Bucle, 3) Then
      Exit Sub
   End If

   r_int_ConVer = 5
   r_int_Cont = 0

   Do While Not r_rst_Bucle.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT CASE WHEN C.DATGEN_RAZSOC IS NULL THEN 'RECURSOS PROPIOS' ELSE TRIM(C.DATGEN_RAZSOC) END AS PROMOTORA,"
      g_str_Parame = g_str_Parame & "        CASE WHEN B.DATGEN_TITULO IS NULL THEN 'BIEN TERMINADO'   ELSE TRIM(B.DATGEN_TITULO) END AS PROYECTO,"
      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN A.HIPMAE_CODPRD END) AS PROD_MIVIVIENDA,"
      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN A.HIPMAE_CODPRD END) AS PROD_MICASITA,"
      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrCOF & ")  THEN A.HIPMAE_CODPRD END) AS PROD_COFICASA,"
      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN A.HIPMAE_CODPRD END) + "
      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN A.HIPMAE_CODPRD END) + "
      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrCOF & ")  THEN A.HIPMAE_CODPRD END) TOTAL_PRODUCTO,"
      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN A.HIPMAE_MTOPRE END) AS MONTO_MIVIVIENDA,"
      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN A.HIPMAE_MTOPRE END) AS MONTO_MICASITA,"
      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrCOF & ")  THEN A.HIPMAE_MTOPRE END) AS MONTO_COFICASA,"
      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN A.HIPMAE_MTOPRE ELSE 0 END) + "
      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN A.HIPMAE_MTOPRE ELSE 0 END) + "
      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrCOF & ")  THEN A.HIPMAE_MTOPRE ELSE 0 END) AS TOTAL_MONTO"
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
      g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.HIPMAE_PRYINM "
      g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = B.DATGEN_VENTDO AND C.DATGEN_EMPNDO = B.DATGEN_VENNDO "
      g_str_Parame = g_str_Parame & " WHERE A.HIPMAE_SITUAC IN (2,6,9) AND A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "  AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
      If r_rst_Bucle!PROMOTORA <> "RECURSOS PROPIOS" Then
         g_str_Parame = g_str_Parame & " AND C.DATGEN_RAZSOC='" & r_rst_Bucle!PROMOTORA & "'  "
      Else
         g_str_Parame = g_str_Parame & " AND C.DATGEN_RAZSOC IS NULL  "
      End If
      g_str_Parame = g_str_Parame & " GROUP BY C.DATGEN_RAZSOC, B.DATGEN_TITULO"
      g_str_Parame = g_str_Parame & " ORDER BY PROMOTORA, PROYECTO, TOTAL_MONTO"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If

      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont + 1
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = g_rst_Princi!PROMOTORA
            
            r_str_Nombre = r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2)
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = g_rst_Princi!PROYECTO
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = g_rst_Princi!PROD_MIVIVIENDA
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = g_rst_Princi!PROD_MICASITA
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = g_rst_Princi!PROD_COFICASA
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = g_rst_Princi!TOTAL_PRODUCTO
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = (g_rst_Princi!PROD_MIVIVIENDA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_MICASITA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_COFICASA / r_int_Porcent) * 100
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = IIf(IsNull(g_rst_Princi!MONTO_MIVIVIENDA), 0, Format(g_rst_Princi!MONTO_MIVIVIENDA, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = IIf(IsNull(g_rst_Princi!MONTO_MICASITA), 0, Format(g_rst_Princi!MONTO_MICASITA, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = IIf(IsNull(g_rst_Princi!MONTO_COFICASA), 0, Format(g_rst_Princi!MONTO_COFICASA, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = IIf(IsNull(g_rst_Princi!TOTAL_MONTO), 0, Format(g_rst_Princi!TOTAL_MONTO, "###,###,##0.00"))

            r_int_Cantidad = r_int_Cantidad + g_rst_Princi!TOTAL_PRODUCTO
            r_obj_Excel.ActiveSheet.Range("H" & r_int_ConVer & ":H" & r_int_ConVer).Select
            r_obj_Excel.ActiveSheet.Range("H" & r_int_ConVer & ":H" & r_int_ConVer).NumberFormat = "#0.#00"
            r_int_Total = r_int_Total + g_rst_Princi!TOTAL_PRODUCTO
            
            If IsNull(g_rst_Princi!MONTO_MIVIVIENDA) Or g_rst_Princi!MONTO_MIVIVIENDA = 0 Then
               r_obj_Excel.ActiveSheet.Range("I" & r_int_ConVer & ":I" & r_int_ConVer).Select
               r_obj_Excel.ActiveSheet.Range("I" & r_int_ConVer & ":I" & r_int_ConVer).NumberFormat = "#0.#00"
            End If

            If IsNull(g_rst_Princi!MONTO_MICASITA) Or g_rst_Princi!MONTO_MICASITA = 0 Then
               r_obj_Excel.ActiveSheet.Range("J" & r_int_ConVer & ":J" & r_int_ConVer).Select
               r_obj_Excel.ActiveSheet.Range("J" & r_int_ConVer & ":J" & r_int_ConVer).NumberFormat = "#0.#00"
            End If
            
            If IsNull(g_rst_Princi!MONTO_COFICASA) Or g_rst_Princi!MONTO_COFICASA = 0 Then
               r_obj_Excel.ActiveSheet.Range("K" & r_int_ConVer & ":K" & r_int_ConVer).Select
               r_obj_Excel.ActiveSheet.Range("K" & r_int_ConVer & ":K" & r_int_ConVer).NumberFormat = "#0.#00"
            End If
            
            r_int_ConVer = r_int_ConVer + 1
            r_dbl_TotGrpPor = r_dbl_TotGrpPor + (g_rst_Princi!PROD_MIVIVIENDA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_MICASITA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_COFICASA / r_int_Porcent) * 100
            r_dbl_TotGrp = r_dbl_TotGrp + (g_rst_Princi!PROD_MIVIVIENDA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_MICASITA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_COFICASA / r_int_Porcent) * 100
            r_int_ProdMV = r_int_ProdMV + g_rst_Princi!PROD_MIVIVIENDA
            r_int_ProdMC = r_int_ProdMC + g_rst_Princi!PROD_MICASITA
            r_int_ProdCC = r_int_ProdCC + g_rst_Princi!PROD_COFICASA
            r_int_TotProd = r_int_ProdMV + r_int_ProdMC + r_int_ProdCC
            
            r_dbl_MontoMV = r_dbl_MontoMV + IIf(IsNull(g_rst_Princi!MONTO_MIVIVIENDA), 0, g_rst_Princi!MONTO_MIVIVIENDA)
            r_dbl_MontoMC = r_dbl_MontoMC + IIf(IsNull(g_rst_Princi!MONTO_MICASITA), 0, g_rst_Princi!MONTO_MICASITA)
            r_dbl_MontoCC = r_dbl_MontoCC + IIf(IsNull(g_rst_Princi!MONTO_COFICASA), 0, g_rst_Princi!MONTO_COFICASA)
            r_dbl_MontoTot = r_dbl_MontoMV + r_dbl_MontoMC + r_dbl_MontoCC
            r_int_TotalMV = r_int_TotalMV + g_rst_Princi!PROD_MIVIVIENDA
            r_int_TotalMC = r_int_TotalMC + g_rst_Princi!PROD_MICASITA
            r_int_TotalCC = r_int_TotalCC + g_rst_Princi!PROD_COFICASA
            r_dbl_MonTotMV = r_dbl_MonTotMV + IIf(IsNull(g_rst_Princi!MONTO_MIVIVIENDA), 0, g_rst_Princi!MONTO_MIVIVIENDA)
            r_dbl_MonTotMC = r_dbl_MonTotMC + IIf(IsNull(g_rst_Princi!MONTO_MICASITA), 0, g_rst_Princi!MONTO_MICASITA)
            r_dbl_MonTotCC = r_dbl_MonTotCC + IIf(IsNull(g_rst_Princi!MONTO_COFICASA), 0, g_rst_Princi!MONTO_COFICASA)
            
            g_rst_Princi.MoveNext
            If Not g_rst_Princi.EOF Then

               If r_str_Nombre <> g_rst_Princi!PROMOTORA Then
                  r_int_Cont = r_int_Cont + 1
                  With r_obj_Excel.ActiveSheet
                     .Cells(r_int_ConVer - 1, 1) = r_int_Cont
                     .Cells(r_int_ConVer, 4) = r_int_ProdMV
                     .Cells(r_int_ConVer, 5) = r_int_ProdMC
                     .Cells(r_int_ConVer, 6) = r_int_ProdCC
                     .Cells(r_int_ConVer, 7) = r_int_TotProd
                     .Cells(r_int_ConVer, 8) = r_dbl_TotGrpPor
                     .Cells(r_int_ConVer, 9) = Format(r_dbl_MontoMV, "###,###,##0.00")
                     .Cells(r_int_ConVer, 10) = Format(r_dbl_MontoMC, "###,###,##0.00")
                     .Cells(r_int_ConVer, 11) = Format(r_dbl_MontoCC, "###,###,##0.00")
                     .Cells(r_int_ConVer, 12) = Format(r_dbl_MontoTot, "###,###,##0.00")
                     
                     If r_dbl_MontoMV = 0 Then
                        .Range("H" & r_int_ConVer & ":H" & r_int_ConVer).Select
                        .Range("H" & r_int_ConVer & ":H" & r_int_ConVer).NumberFormat = "#0.#00"
                     End If
                     If r_dbl_MontoMC = 0 Then
                        .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).Select
                        .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).NumberFormat = "#0.#00"
                     End If

                     .Range(.Cells(r_int_ConVer, 3), .Cells(r_int_ConVer, 3)).HorizontalAlignment = xlHAlignRight
                     .Cells(r_int_ConVer, 3) = "Total Agrupado : "
                     .Range("G" & r_int_ConVer & ":G" & r_int_ConVer).Select
                     .Range("G" & r_int_ConVer & ":G" & r_int_ConVer).NumberFormat = "#0.#00"
                     .Range("A" & r_int_ConVer & ":J" & r_int_ConVer).Interior.Color = RGB(146, 208, 80)
                     .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 12)).Font.Bold = True
                  End With

                  r_int_Cantidad = 0
                  r_int_ProdMV = 0
                  r_int_ProdMC = 0
                  r_int_ProdCC = 0
                  r_dbl_MontoMV = 0
                  r_dbl_MontoMC = 0
                  r_dbl_MontoCC = 0
                  r_dbl_TotGrpPor = 0
                  r_int_ConVer = r_int_ConVer + 1
               End If
            Else
               r_int_Cont = r_int_Cont + 1
               
               'Totalizadores
               With r_obj_Excel.ActiveSheet
                  .Cells(r_int_ConVer, 4) = r_int_ProdMV
                  .Cells(r_int_ConVer, 5) = r_int_ProdMC
                  .Cells(r_int_ConVer, 6) = r_int_ProdCC
                  .Cells(r_int_ConVer, 7) = r_int_TotProd
                  .Cells(r_int_ConVer, 8) = r_dbl_TotGrpPor
                  .Cells(r_int_ConVer, 9) = Format(r_dbl_MontoMV, "###,###,##0.00")
                  .Cells(r_int_ConVer, 10) = Format(r_dbl_MontoMC, "###,###,##0.00")
                  .Cells(r_int_ConVer, 11) = Format(r_dbl_MontoCC, "###,###,##0.00")
                  .Cells(r_int_ConVer, 12) = Format(r_dbl_MontoTot, "###,###,##0.00")

                  .Range("H" & r_int_ConVer & ":H" & r_int_ConVer).Select
                  .Range("H" & r_int_ConVer & ":H" & r_int_ConVer).NumberFormat = "#0.#00"
                  .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).Select
                  .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).NumberFormat = "###,###,##0.00"
                  .Range("J" & r_int_ConVer & ":J" & r_int_ConVer).Select
                  .Range("J" & r_int_ConVer & ":J" & r_int_ConVer).NumberFormat = "###,###,##0.00"
                  .Range("K" & r_int_ConVer & ":K" & r_int_ConVer).Select
                  .Range("K" & r_int_ConVer & ":K" & r_int_ConVer).NumberFormat = "###,###,##0.00"

                  .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 12)).Font.Bold = True
               End With

               r_int_ProdMV = 0
               r_int_ProdMC = 0
               r_int_ProdCC = 0
               r_dbl_MontoMV = 0
               r_dbl_MontoMC = 0
               r_dbl_MontoCC = 0
               r_dbl_TotGrpPor = 0
               r_int_ConVer = r_int_ConVer + 1
            End If
         Loop

         With r_obj_Excel.ActiveSheet
            .Range(.Cells(r_int_ConVer - 1, 3), .Cells(r_int_ConVer - 1, 3)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer - 1, 3) = "Total Agrupado : "
            .Range("A" & r_int_ConVer - 1 & ":L" & r_int_ConVer - 1).Interior.Color = RGB(146, 208, 80)
         End With
      End If

      r_rst_Bucle.MoveNext
   Loop

   r_int_Cont1 = 3
   r_int_ConVer1 = 5

   Do While r_int_Cont1 < r_int_ConVer + 1
      r_obj_Excel.ActiveSheet.Range("A5:A" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("B5:B" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("B5:B" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(3, 1), r_obj_Excel.Cells(3, 12)).Select
      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("C3:C" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("D4:F" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("H4:I" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("I4:K" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      
      If r_int_ConVer1 > 4 Then
         If (r_int_ConVer1 - 2) = 1 Then
            r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 - 2, 1), r_obj_Excel.Cells(r_int_ConVer1 - 2, 12)).Select
            r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
         End If
         r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1, 1), r_obj_Excel.Cells(r_int_ConVer1, 12)).Select
         r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      End If

      r_obj_Excel.ActiveSheet.Range("D3:D" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("E4:E" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("F4:F" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("G3:G" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("H3:H" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("I3:I" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("J4:J" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("K4:K" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("L3:L" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("M3:M" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_int_Cont1 = r_int_Cont1 + 1
      r_int_ConVer1 = r_int_ConVer1 + 1
   Loop

   With r_obj_Excel.ActiveSheet
      .Cells(r_int_ConVer1 - 2, 3) = "TOTALES"
      .Cells(r_int_ConVer1 - 2, 4) = r_int_TotalMV
      .Cells(r_int_ConVer1 - 2, 5) = r_int_TotalMC
      .Cells(r_int_ConVer1 - 2, 6) = r_int_TotalCC
      .Cells(r_int_ConVer1 - 2, 7) = r_int_TotalMV + r_int_TotalMC + r_int_TotalCC

      .Range("H" & r_int_ConVer1 - 2 & ":H" & r_int_ConVer1 - 2).Select
      .Range("H" & r_int_ConVer1 - 2 & ":H" & r_int_ConVer1 - 2).NumberFormat = "#0.#00"
      .Cells(r_int_ConVer1 - 2, 8) = r_dbl_TotGrp

      .Cells(r_int_ConVer1 - 2, 9) = Format(r_dbl_MonTotMV, "###,###,##0.00")
      .Cells(r_int_ConVer1 - 2, 10) = Format(r_dbl_MonTotMC, "###,###,##0.00")
      .Cells(r_int_ConVer1 - 2, 11) = Format(r_dbl_MonTotCC, "###,###,##0.00")
      .Cells(r_int_ConVer1 - 2, 12) = Format(r_dbl_MonTotMV + r_dbl_MonTotMC + r_dbl_MonTotCC, "###,###,##0.00")


      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer1 - 2, 3), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer1 - 2, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(r_int_ConVer + 1, 3), .Cells(r_int_ConVer + 1, 12)).Font.Bold = True

      .Range("A" & r_int_ConVer + 1 & ":L" & r_int_ConVer + 1).Interior.Color = RGB(239, 215, 155)
      .Range("A" & 3 & ":L" & 3).Interior.Color = RGB(213, 239, 245)
      .Range("A" & 4 & ":L" & 4).Interior.Color = RGB(213, 239, 245)

      .Range("A1:A2").Select
   End With

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub


'Private Sub fs_GenExc_Promotora()
'Dim r_rst_Bucle      As ADODB.Recordset
'Dim r_obj_Excel      As Excel.Application
'Dim r_int_ConVer     As Integer
'Dim r_int_Cont       As Integer
'Dim r_int_Cont1      As Integer
'Dim r_int_ConVer1    As Integer
'Dim r_str_Nombre     As String
'Dim r_int_ProdMV     As Integer
'Dim r_int_ProdMC     As Integer
'Dim r_int_TotProd    As Integer
'Dim r_dbl_MontoMV    As Double
'Dim r_dbl_MontoMC    As Double
'Dim r_dbl_MontoTot   As Double
'Dim r_int_TotalMV    As Integer
'Dim r_int_TotalMC    As Integer
'Dim r_dbl_MonTotMV   As Double
'Dim r_dbl_MonTotMC   As Double
'Dim r_int_Porcent    As Integer
'Dim r_int_Cantidad   As Integer
'Dim r_int_Total      As Integer
'Dim r_dbl_TotGrpPor  As Double
'Dim r_dbl_TotGrp     As Double
'
'   Set r_obj_Excel = New Excel.Application
'   r_obj_Excel.SheetsInNewWorkbook = 1
'   r_obj_Excel.Workbooks.Add
'
'   With r_obj_Excel.ActiveSheet
'      .Cells(1, 1) = "DETALLE DE DESEMBOLSOS POR PROMOTORAS DEL " & CStr(ipp_FecIni.Text) & " AL " & CStr((ipp_FecFin.Text))
'      .Range("A1:J1").Select
'      .Range("A1:J1").HorizontalAlignment = xlHAlignCenter
'      .Range("A1:J1").Font.Bold = True
'      r_obj_Excel.Selection.MergeCells = True
'
'      .Cells(3, 1) = "ITEM"
'      .Cells(3, 2) = "PROMOTORAS"
'      .Cells(3, 3) = "PROYECTOS"
'      .Cells(4, 4) = "MiVivienda"
'      .Cells(4, 5) = "MiCasita"
'      .Cells(4, 6) = "Coficasa"
'      .Cells(3, 5) = "PRODUCTO"
'      .Cells(4, 7) = "TOTAL POR PRODUCTO"
'      .Cells(4, 8) = "%"
'      .Cells(4, 9) = "MiVivienda"
'      .Cells(4, 10) = "MiCasita"
'      .Cells(3, 9) = "MONTO"
'      .Cells(4, 11) = "TOTAL POR MONTO"
'
'      'r_obj_Excel.Visible = True
'      .Range("F3:F4").Select
'      '.Range("G3:G4").Select
'      r_obj_Excel.Selection.MergeCells = True
'      r_obj_Excel.Selection.Cells.WrapText = True
'
'      .Range("G3:G4").Select
'      r_obj_Excel.Selection.MergeCells = True
'      r_obj_Excel.Selection.Cells.WrapText = True
'
'      .Range("D3:F3").Select
'      r_obj_Excel.Selection.MergeCells = True
'
'      .Range("J3:J4").Select
'      r_obj_Excel.Selection.MergeCells = True
'      r_obj_Excel.Selection.Cells.WrapText = True
'
'      .Range("H3:I3").Select
'      r_obj_Excel.Selection.MergeCells = True
'
'      .Range("A3:A4").Select
'      r_obj_Excel.Selection.MergeCells = True
'
'      .Range("B3:B4").Select
'      r_obj_Excel.Selection.MergeCells = True
'
'      .Range("C3:C4").Select
'      r_obj_Excel.Selection.MergeCells = True
'
'      .Range(.Cells(3, 1), .Cells(3, 12)).Font.Bold = True
'      .Range(.Cells(3, 1), .Cells(3, 12)).HorizontalAlignment = xlHAlignCenter
'      .Range(.Cells(4, 1), .Cells(4, 12)).Font.Bold = True
'      .Range(.Cells(4, 1), .Cells(4, 12)).HorizontalAlignment = xlHAlignCenter
'
'      .Columns("A").ColumnWidth = 6
'      .Columns("A").HorizontalAlignment = xlHAlignCenter
'      .Columns("B").ColumnWidth = 50
'      .Columns("C").ColumnWidth = 50
'      .Columns("D").ColumnWidth = 11
'      .Columns("D").HorizontalAlignment = xlHAlignCenter
'      .Columns("E").ColumnWidth = 11
'      .Columns("E").HorizontalAlignment = xlHAlignCenter
'      .Columns("F").ColumnWidth = 11
'      .Columns("F").HorizontalAlignment = xlHAlignCenter
'      .Columns("G").ColumnWidth = 8
'      .Columns("H").ColumnWidth = 17
'      .Columns("I").ColumnWidth = 17
'      .Columns("J").ColumnWidth = 17
'   End With
'
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT COUNT(A.HIPMAE_CODPRD) AS CONTGRUPO "
'   g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A "
'   g_str_Parame = g_str_Parame & "   LEFT JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.HIPMAE_PRYINM "
'   g_str_Parame = g_str_Parame & "   LEFT JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = B.DATGEN_VENTDO AND C.DATGEN_EMPNDO = B.DATGEN_VENNDO "
'   g_str_Parame = g_str_Parame & "  WHERE A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "  AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'       Exit Sub
'   End If
'
'   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
'      r_int_Porcent = g_rst_Princi!CONTGRUPO
'   End If
'
'   'Para ser tomado por consejero en el proximo query
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "SELECT CASE WHEN C.DATGEN_RAZSOC IS NULL THEN 'RECURSOS PROPIOS' ELSE TRIM(C.DATGEN_RAZSOC) END AS PROMOTORA, "
'   g_str_Parame = g_str_Parame & "       COUNT(CASE WHEN A.HIPMAE_CODPRD <> '011' THEN A.HIPMAE_CODPRD END) + COUNT(CASE WHEN A.HIPMAE_CODPRD =  '011' THEN A.HIPMAE_CODPRD END) TOTAL_PRODUCTO "
'   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
'   g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.HIPMAE_PRYINM "
'   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = B.DATGEN_VENTDO AND C.DATGEN_EMPNDO = B.DATGEN_VENNDO "
'   g_str_Parame = g_str_Parame & " WHERE A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "  AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
'   g_str_Parame = g_str_Parame & " GROUP BY C.DATGEN_RAZSOC ORDER BY 2 DESC"
'
'   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Bucle, 3) Then
'      Exit Sub
'   End If
'
'   r_int_ConVer = 5
'   r_int_Cont = 0
'
'   Do While Not r_rst_Bucle.EOF
'      g_str_Parame = ""
'      g_str_Parame = g_str_Parame & " SELECT CASE WHEN C.DATGEN_RAZSOC IS NULL THEN 'RECURSOS PROPIOS' ELSE TRIM(C.DATGEN_RAZSOC) END AS PROMOTORA,"
'      g_str_Parame = g_str_Parame & "        CASE WHEN B.DATGEN_TITULO IS NULL THEN 'BIEN TERMINADO'   ELSE TRIM(B.DATGEN_TITULO) END AS PROYECTO,"
'      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN (A.HIPMAE_CODPRD <> '011' AND A.HIPMAE_CODPRD <> '002' AND A.HIPMAE_CODPRD <> '006' AND A.HIPMAE_CODPRD <> '020') THEN A.HIPMAE_CODPRD END) AS PROD_MIVIVIENDA,"
'      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN (A.HIPMAE_CODPRD =  '011' AND A.HIPMAE_CODPRD =  '002' AND A.HIPMAE_CODPRD =  '006')  THEN A.HIPMAE_CODPRD END) AS PROD_MICASITA,"
'      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN (A.HIPMAE_CODPRD =  '020') THEN A.HIPMAE_CODPRD END) AS PROD_COFICASA,"
'      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN (A.HIPMAE_CODPRD <> '011' AND A.HIPMAE_CODPRD <> '002' AND A.HIPMAE_CODPRD <> '006' AND A.HIPMAE_CODPRD <> '020') THEN A.HIPMAE_CODPRD END) + "
'      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN (A.HIPMAE_CODPRD =  '011' AND A.HIPMAE_CODPRD =  '002' AND A.HIPMAE_CODPRD =  '006') THEN A.HIPMAE_CODPRD END) + "
'      g_str_Parame = g_str_Parame & "        COUNT(CASE WHEN (A.HIPMAE_CODPRD =  '020') THEN A.HIPMAE_CODPRD END) TOTAL_PRODUCTO,"
'      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN (A.HIPMAE_CODPRD <> '011' AND A.HIPMAE_CODPRD <> '002' AND A.HIPMAE_CODPRD <> '006' AND A.HIPMAE_CODPRD <> '020') THEN A.HIPMAE_MTOPRE END) AS MONTO_MIVIVIENDA,"
'      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN (A.HIPMAE_CODPRD =  '011' AND A.HIPMAE_CODPRD =  '002' AND A.HIPMAE_CODPRD =  '006') THEN A.HIPMAE_MTOPRE END) AS MONTO_MICASITA,"
'      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN (A.HIPMAE_CODPRD =  '020') THEN A.HIPMAE_MTOPRE END) AS MONTO_COFICASA,"
'      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN (A.HIPMAE_CODPRD <> '011' AND A.HIPMAE_CODPRD <> '002' AND A.HIPMAE_CODPRD <> '006' AND A.HIPMAE_CODPRD <> '020') THEN A.HIPMAE_MTOPRE ELSE 0 END) + "
'      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN (A.HIPMAE_CODPRD =  '011' AND A.HIPMAE_CODPRD =  '002' AND A.HIPMAE_CODPRD =  '006') THEN A.HIPMAE_MTOPRE ELSE 0 END) + "
'      g_str_Parame = g_str_Parame & "          SUM(CASE WHEN (A.HIPMAE_CODPRD =  '020') THEN A.HIPMAE_MTOPRE ELSE 0 END) AS TOTAL_MONTO"
'      g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
'      g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.HIPMAE_PRYINM "
'      g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = B.DATGEN_VENTDO AND C.DATGEN_EMPNDO = B.DATGEN_VENNDO "
'      g_str_Parame = g_str_Parame & " WHERE A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "  AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
'      If r_rst_Bucle!PROMOTORA <> "RECURSOS PROPIOS" Then
'         g_str_Parame = g_str_Parame & " AND C.DATGEN_RAZSOC='" & r_rst_Bucle!PROMOTORA & "'  "
'      Else
'         g_str_Parame = g_str_Parame & " AND C.DATGEN_RAZSOC IS NULL  "
'      End If
'      g_str_Parame = g_str_Parame & " GROUP BY C.DATGEN_RAZSOC, B.DATGEN_TITULO"
'      g_str_Parame = g_str_Parame & " ORDER BY PROMOTORA, PROYECTO, TOTAL_MONTO"
'
'      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'          Exit Sub
'      End If
'
'      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
'         g_rst_Princi.MoveFirst
'         Do While Not g_rst_Princi.EOF
'            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont + 1
'            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = g_rst_Princi!PROMOTORA
'            r_str_Nombre = r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) 'tiene q estar aqui para poder tomar el nombre de la promotora
'            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = g_rst_Princi!PROYECTO
'            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = g_rst_Princi!PROD_MIVIVIENDA
'            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = g_rst_Princi!PROD_MICASITA
'            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = g_rst_Princi!TOTAL_PRODUCTO
'            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = (g_rst_Princi!PROD_MIVIVIENDA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_MICASITA / r_int_Porcent) * 100
'            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = IIf(IsNull(g_rst_Princi!MONTO_MIVIVIENDA), 0, Format(g_rst_Princi!MONTO_MIVIVIENDA, "###,###,##0.00"))
'            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = IIf(IsNull(g_rst_Princi!MONTO_MICASITA), 0, Format(g_rst_Princi!MONTO_MICASITA, "###,###,##0.00"))
'            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = IIf(IsNull(g_rst_Princi!TOTAL_MONTO), 0, Format(g_rst_Princi!TOTAL_MONTO, "###,###,##0.00"))
'
'            r_int_Cantidad = r_int_Cantidad + g_rst_Princi!TOTAL_PRODUCTO
'            r_obj_Excel.ActiveSheet.Range("G" & r_int_ConVer & ":G" & r_int_ConVer).Select
'            r_obj_Excel.ActiveSheet.Range("G" & r_int_ConVer & ":G" & r_int_ConVer).NumberFormat = "#0.#00"
'            r_int_Total = r_int_Total + g_rst_Princi!TOTAL_PRODUCTO
'            If IsNull(g_rst_Princi!MONTO_MIVIVIENDA) Or g_rst_Princi!MONTO_MIVIVIENDA = 0 Then
'               r_obj_Excel.ActiveSheet.Range("H" & r_int_ConVer & ":H" & r_int_ConVer).Select
'               r_obj_Excel.ActiveSheet.Range("H" & r_int_ConVer & ":H" & r_int_ConVer).NumberFormat = "#0.#00"
'            End If
'
'            If IsNull(g_rst_Princi!MONTO_MICASITA) Or g_rst_Princi!MONTO_MICASITA = 0 Then
'               r_obj_Excel.ActiveSheet.Range("I" & r_int_ConVer & ":I" & r_int_ConVer).Select
'               r_obj_Excel.ActiveSheet.Range("I" & r_int_ConVer & ":I" & r_int_ConVer).NumberFormat = "#0.#00"
'            End If
'
'            r_int_ConVer = r_int_ConVer + 1
'            r_dbl_TotGrpPor = r_dbl_TotGrpPor + (g_rst_Princi!PROD_MIVIVIENDA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_MICASITA / r_int_Porcent) * 100
'            r_dbl_TotGrp = r_dbl_TotGrp + (g_rst_Princi!PROD_MIVIVIENDA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_MICASITA / r_int_Porcent) * 100
'            r_int_ProdMV = r_int_ProdMV + g_rst_Princi!PROD_MIVIVIENDA
'            r_int_ProdMC = r_int_ProdMC + g_rst_Princi!PROD_MICASITA
'            r_int_TotProd = r_int_ProdMV + r_int_ProdMC
'
'            r_dbl_MontoMV = r_dbl_MontoMV + IIf(IsNull(g_rst_Princi!MONTO_MIVIVIENDA), 0, g_rst_Princi!MONTO_MIVIVIENDA)
'            r_dbl_MontoMC = r_dbl_MontoMC + IIf(IsNull(g_rst_Princi!MONTO_MICASITA), 0, g_rst_Princi!MONTO_MICASITA)
'            r_dbl_MontoTot = r_dbl_MontoMV + r_dbl_MontoMC
'            r_int_TotalMV = r_int_TotalMV + g_rst_Princi!PROD_MIVIVIENDA
'            r_int_TotalMC = r_int_TotalMC + g_rst_Princi!PROD_MICASITA
'            r_dbl_MonTotMV = r_dbl_MonTotMV + IIf(IsNull(g_rst_Princi!MONTO_MIVIVIENDA), 0, g_rst_Princi!MONTO_MIVIVIENDA)
'            r_dbl_MonTotMC = r_dbl_MonTotMC + IIf(IsNull(g_rst_Princi!MONTO_MICASITA), 0, g_rst_Princi!MONTO_MICASITA)
'
'            g_rst_Princi.MoveNext
'            If Not g_rst_Princi.EOF Then
'
'               If r_str_Nombre <> g_rst_Princi!PROMOTORA Then
'                  r_int_Cont = r_int_Cont + 1
'
'                  With r_obj_Excel.ActiveSheet
'                     .Cells(r_int_ConVer - 1, 1) = r_int_Cont
'                     .Cells(r_int_ConVer, 4) = r_int_ProdMV
'                     .Cells(r_int_ConVer, 5) = r_int_ProdMC
'                     .Cells(r_int_ConVer, 6) = r_int_TotProd
'                     .Cells(r_int_ConVer, 7) = r_dbl_TotGrpPor
'                     .Cells(r_int_ConVer, 8) = Format(r_dbl_MontoMV, "###,###,##0.00")
'                     .Cells(r_int_ConVer, 9) = Format(r_dbl_MontoMC, "###,###,##0.00")
'                     .Cells(r_int_ConVer, 10) = Format(r_dbl_MontoTot, "###,###,##0.00")
'
'                     If r_dbl_MontoMV = 0 Then
'                        .Range("H" & r_int_ConVer & ":H" & r_int_ConVer).Select
'                        .Range("H" & r_int_ConVer & ":H" & r_int_ConVer).NumberFormat = "#0.#00"
'                     End If
'                     If r_dbl_MontoMC = 0 Then
'                        .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).Select
'                        .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).NumberFormat = "#0.#00"
'                     End If
'
'                     .Range(.Cells(r_int_ConVer, 3), .Cells(r_int_ConVer, 3)).HorizontalAlignment = xlHAlignRight
'                     .Cells(r_int_ConVer, 3) = "Total Agrupado : "
'                     .Range("G" & r_int_ConVer & ":G" & r_int_ConVer).Select
'                     .Range("G" & r_int_ConVer & ":G" & r_int_ConVer).NumberFormat = "#0.#00"
'                     .Range("A" & r_int_ConVer & ":J" & r_int_ConVer).Interior.Color = RGB(146, 208, 80)
'                     .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 12)).Font.Bold = True
'                  End With
'
'                  r_int_Cantidad = 0
'                  r_int_ProdMV = 0
'                  r_int_ProdMC = 0
'                  r_dbl_MontoMV = 0
'                  r_dbl_MontoMC = 0
'                  r_dbl_TotGrpPor = 0
'                  r_int_ConVer = r_int_ConVer + 1
'               End If
'
'            Else
'               r_int_Cont = r_int_Cont + 1
'
'               'Totalizadores
'               With r_obj_Excel.ActiveSheet
'                  .Cells(r_int_ConVer, 4) = r_int_ProdMV
'                  .Cells(r_int_ConVer, 5) = r_int_ProdMC
'                  .Cells(r_int_ConVer, 6) = r_int_TotProd
'                  .Cells(r_int_ConVer, 7) = r_dbl_TotGrpPor
'                  .Cells(r_int_ConVer, 8) = Format(r_dbl_MontoMV, "###,###,##0.00")
'                  .Cells(r_int_ConVer, 9) = Format(r_dbl_MontoMC, "###,###,##0.00")
'                  .Cells(r_int_ConVer, 10) = Format(r_dbl_MontoTot, "###,###,##0.00")
'
'                  .Range("G" & r_int_ConVer & ":G" & r_int_ConVer).Select
'                  .Range("G" & r_int_ConVer & ":G" & r_int_ConVer).NumberFormat = "#0.#00"
'                  .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).Select
'                  .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).NumberFormat = "###,###,##0.00"
'                  .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 12)).Font.Bold = True
'               End With
'
'               r_int_ProdMV = 0
'               r_int_ProdMC = 0
'               r_dbl_MontoMV = 0
'               r_dbl_MontoMC = 0
'               r_dbl_TotGrpPor = 0
'               r_int_ConVer = r_int_ConVer + 1
'            End If
'         Loop
'
'         With r_obj_Excel.ActiveSheet
'            .Range(.Cells(r_int_ConVer - 1, 3), .Cells(r_int_ConVer - 1, 3)).HorizontalAlignment = xlHAlignRight
'            .Cells(r_int_ConVer - 1, 3) = "Total Agrupado : "
'            .Range("A" & r_int_ConVer - 1 & ":J" & r_int_ConVer - 1).Interior.Color = RGB(146, 208, 80)
'         End With
'      End If
'
'      r_rst_Bucle.MoveNext
'   Loop
'
'   r_int_Cont1 = 3
'   r_int_ConVer1 = 5
'
'   Do While r_int_Cont1 < r_int_ConVer + 1
'      r_obj_Excel.ActiveSheet.Range("A5:A" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("B5:B" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("B5:B" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'      r_obj_Excel.Range(r_obj_Excel.Cells(3, 1), r_obj_Excel.Cells(3, 10)).Select
'      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("C3:C" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("D4:E" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("H4:I" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'
'      If r_int_ConVer1 > 4 Then
'         If (r_int_ConVer1 - 2) = 1 Then
'            r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 - 2, 1), r_obj_Excel.Cells(r_int_ConVer1 - 2, 10)).Select
'            r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'         End If
'         r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1, 1), r_obj_Excel.Cells(r_int_ConVer1, 10)).Select
'         r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'      End If
'
'      r_obj_Excel.ActiveSheet.Range("D3:D" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("E4:E" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("F4:F" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("G3:G" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("H3:H" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("I4:I" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("J3:J" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'      r_obj_Excel.ActiveSheet.Range("K3:K" & 1 + r_int_Cont1).Select
'      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'      r_int_Cont1 = r_int_Cont1 + 1
'      r_int_ConVer1 = r_int_ConVer1 + 1
'   Loop
'
'   With r_obj_Excel.ActiveSheet
'      .Cells(r_int_ConVer1 - 2, 3) = "TOTALES"
'      .Cells(r_int_ConVer1 - 2, 4) = r_int_TotalMV
'      .Cells(r_int_ConVer1 - 2, 5) = r_int_TotalMC
'      .Cells(r_int_ConVer1 - 2, 6) = r_int_TotalMV + r_int_TotalMC
'      .Cells(r_int_ConVer1 - 2, 8) = Format(r_dbl_MonTotMV, "###,###,##0.00")
'      .Cells(r_int_ConVer1 - 2, 9) = Format(r_dbl_MonTotMC, "###,###,##0.00")
'      .Cells(r_int_ConVer1 - 2, 10) = Format(r_dbl_MonTotMV + r_dbl_MonTotMC, "###,###,##0.00")
'
'      .Range("G" & r_int_ConVer1 - 2 & ":G" & r_int_ConVer1 - 2).Select
'      .Range("G" & r_int_ConVer1 - 2 & ":G" & r_int_ConVer1 - 2).NumberFormat = "#0.#00"
'      .Cells(r_int_ConVer1 - 2, 7) = r_dbl_TotGrp
'
'      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer1 - 2, 3), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer1 - 2, 3)).HorizontalAlignment = xlHAlignCenter
'      .Range(.Cells(r_int_ConVer + 1, 3), .Cells(r_int_ConVer + 1, 10)).Font.Bold = True
'
'      .Range("A" & r_int_ConVer + 1 & ":J" & r_int_ConVer + 1).Interior.Color = RGB(239, 215, 155)
'      .Range("A" & 3 & ":J" & 3).Interior.Color = RGB(213, 239, 245)
'      .Range("A" & 4 & ":J" & 4).Interior.Color = RGB(213, 239, 245)
'
'      .Range("A1:A2").Select
'   End With
'
'   g_rst_Princi.Close
'   Set g_rst_Princi = Nothing
'
'   r_obj_Excel.Visible = True
'   Set r_obj_Excel = Nothing
'End Sub

Private Sub fs_GenExc_Consejero()
Dim r_rst_Bucle      As ADODB.Recordset
Dim r_rst_Otros      As ADODB.Recordset
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_ConVer1    As Integer
Dim r_int_Cont       As Integer
Dim r_int_Cont1      As Integer
Dim r_str_Nombre     As String
Dim r_int_Cantidad   As Integer
Dim r_int_Total      As Integer
Dim s_str_Inicio     As String
Dim s_str_Final      As String
Dim r_int_Porcent    As Integer
Dim r_dbl_TotPorc    As Double
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "DETALLE DE DESEMBOLSOS POR CONSEJEROS DEL " & CStr(ipp_FecIni.Text) & " AL " & CStr((ipp_FecFin.Text))
      .Range("A1:G1").Select
      .Range("A1:G1").HorizontalAlignment = xlHAlignCenter
      .Range("A1:G1").Font.Bold = True
      r_obj_Excel.Selection.MergeCells = True
            
      .Cells(3, 1) = "ITEM"
      .Cells(3, 2) = "CONSEJEROS"
      .Cells(3, 3) = "PROMOTOR"
      .Cells(3, 4) = "PROYECTOS"
      .Cells(3, 5) = "DESEMB."
      .Cells(3, 6) = "TOTAL"
      .Cells(3, 7) = "%"
      .Range(.Cells(3, 1), .Cells(3, 7)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(3, 7)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 6
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 15
      .Columns("C").ColumnWidth = 50
      .Columns("D").ColumnWidth = 50
      .Columns("E").ColumnWidth = 8
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 8
      .Columns("F").HorizontalAlignment = xlHAlignCenter
   End With

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT COUNT(A.HIPMAE_CODPRD) AS CONTGRUPO "
   g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & "   LEFT JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.HIPMAE_PRYINM "
   g_str_Parame = g_str_Parame & "   LEFT JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = B.DATGEN_VENTDO AND C.DATGEN_EMPNDO = B.DATGEN_VENNDO "
   g_str_Parame = g_str_Parame & "  WHERE A.HIPMAE_SITUAC IN (2,6,9) AND A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "  AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_Porcent = g_rst_Princi!CONTGRUPO
   End If

   'Para ser tomado por consejero en el proximo query
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(A.HIPMAE_CONHIP) AS CONSEJERO, COUNT(A.HIPMAE_CODPRD) AS CONTGRUPO "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.HIPMAE_PRYINM "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = B.DATGEN_VENTDO AND C.DATGEN_EMPNDO = B.DATGEN_VENNDO "
   g_str_Parame = g_str_Parame & " WHERE A.HIPMAE_SITUAC IN (2,6,9) AND A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "  AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & " GROUP BY A.HIPMAE_CONHIP ORDER BY 2 DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Bucle, 3) Then
      Exit Sub
   End If
   
   r_int_ConVer = 4
   r_int_Cont = 0
   r_int_Cont1 = 1
   Do While Not r_rst_Bucle.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT TRIM(A.HIPMAE_CONHIP) AS CONSEJERO,"
      g_str_Parame = g_str_Parame & "        TRIM(EJECMC_APEPAT) || ' ' || TRIM(EJECMC_NOMBRE) AS NOMBCONSEJERO,"
      g_str_Parame = g_str_Parame & "        CASE WHEN C.DATGEN_RAZSOC IS NULL THEN 'RECURSOS PROPIOS' ELSE TRIM(C.DATGEN_RAZSOC) END AS PROMOTORA,"
      g_str_Parame = g_str_Parame & "        CASE WHEN B.DATGEN_TITULO IS NULL THEN 'BIEN TERMINADO'   ELSE TRIM(B.DATGEN_TITULO) END AS PROYECTO,"
      g_str_Parame = g_str_Parame & "        COUNT(A.HIPMAE_CODPRD) AS CONTGRUPO"
      g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A"
      g_str_Parame = g_str_Parame & "   LEFT JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.HIPMAE_PRYINM"
      g_str_Parame = g_str_Parame & "   LEFT JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = B.DATGEN_VENTDO AND C.DATGEN_EMPNDO = B.DATGEN_VENNDO"
      g_str_Parame = g_str_Parame & "   LEFT JOIN CRE_EJECMC D ON A.HIPMAE_CONHIP = D.EJECMC_CODEJE "
      g_str_Parame = g_str_Parame & "  WHERE A.HIPMAE_SITUAC IN (2,6,9) AND A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "  AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
      g_str_Parame = g_str_Parame & "    AND A.HIPMAE_CONHIP = '" & r_rst_Bucle!CONSEJERO & "'"
      g_str_Parame = g_str_Parame & "  GROUP BY A.HIPMAE_CONHIP, TRIM(EJECMC_APEPAT) || ' ' || TRIM(EJECMC_NOMBRE), C.DATGEN_RAZSOC, B.DATGEN_TITULO"
      g_str_Parame = g_str_Parame & "  ORDER BY CONSEJERO, PROMOTORA, PROYECTO"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         
         Do While Not g_rst_Princi.EOF
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont + 1
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = g_rst_Princi!NOMBCONSEJERO
            r_obj_Excel.ActiveSheet.Range("B5:B" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_str_Nombre = r_rst_Bucle!CONSEJERO
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = g_rst_Princi!PROMOTORA
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = g_rst_Princi!PROYECTO
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = g_rst_Princi!CONTGRUPO
            
            r_int_Cantidad = r_int_Cantidad + g_rst_Princi!CONTGRUPO
            r_int_Total = r_int_Total + g_rst_Princi!CONTGRUPO
            
            g_rst_Princi.MoveNext
            If Not g_rst_Princi.EOF Then
               
               If r_str_Nombre <> g_rst_Princi!CONSEJERO Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = r_int_Cantidad
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = (r_int_Cantidad / r_int_Porcent) * 100
                  r_dbl_TotPorc = r_dbl_TotPorc + (r_int_Cantidad / r_int_Porcent) * 100
                  r_int_Cont = r_int_Cont + 1
                  r_int_Cantidad = 0
                  s_str_Final = r_int_ConVer
                  
                  If s_str_Inicio <> "" Then
                     With r_obj_Excel.ActiveSheet
                        .Range("A" & s_str_Inicio & ":A" & r_int_ConVer).Select
                        r_obj_Excel.Selection.MergeCells = True
                        .Range("B" & s_str_Inicio & ":B" & r_int_ConVer).Select
                        r_obj_Excel.Selection.MergeCells = True
                        .Range("F" & s_str_Inicio & ":F" & r_int_ConVer).Select
                        r_obj_Excel.Selection.MergeCells = True
                        .Range("G" & s_str_Inicio & ":G" & r_int_ConVer).Select
                        r_obj_Excel.Selection.MergeCells = True
                        .Range("G" & s_str_Inicio & ":G" & r_int_ConVer).NumberFormat = "#0.#00"
                        s_str_Inicio = r_int_ConVer + 2
                     End With
                  End If
                  
                  r_int_ConVer = r_int_ConVer + 1
                  r_obj_Excel.ActiveSheet.Range("A" & r_int_ConVer & ":G" & r_int_ConVer).Interior.Color = RGB(146, 208, 80)
               Else
                  
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = ""
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = ""
                  r_int_Cont1 = r_int_Cont1 + 1
                  If s_str_Inicio = "" Then s_str_Inicio = r_int_ConVer
               End If
            Else
               r_int_Cont = r_int_Cont + 1
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = r_int_Cantidad
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = (r_int_Cantidad / r_int_Porcent) * 100
               r_dbl_TotPorc = r_dbl_TotPorc + (r_int_Cantidad / r_int_Porcent) * 100
                           
               s_str_Final = r_int_ConVer
               r_int_Cantidad = 0
               
               If s_str_Inicio <> "" Then
                  With r_obj_Excel.ActiveSheet
                     .Range("A" & s_str_Inicio & ":A" & r_int_ConVer).Select
                     r_obj_Excel.Selection.MergeCells = True
                     .Range("B" & s_str_Inicio & ":B" & r_int_ConVer).Select
                     r_obj_Excel.Selection.MergeCells = True
                     .Range("F" & s_str_Inicio & ":F" & r_int_ConVer).Select
                     r_obj_Excel.Selection.MergeCells = True
                     .Range("G" & s_str_Inicio & ":G" & r_int_ConVer).Select
                     r_obj_Excel.Selection.MergeCells = True
                     .Range("G" & s_str_Inicio & ":G" & r_int_ConVer).NumberFormat = "#0.#00"
                     s_str_Inicio = r_int_ConVer + 2
                  End With
               End If
               
               r_int_Cont1 = 1
               r_int_ConVer = r_int_ConVer + 1
               r_obj_Excel.ActiveSheet.Range("A" & r_int_ConVer & ":G" & r_int_ConVer).Interior.Color = RGB(146, 208, 80)
            End If
            
            r_int_ConVer = r_int_ConVer + 1
         Loop
      
      End If
      r_rst_Bucle.MoveNext
   Loop
   
   'Para mostrar aquellos consejeros que no tienen movimientos hipotecarios
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT TRIM(EJETIP_CODEJE) EJETIP_CODEJE, TRIM(EJECMC_APEPAT) || ' ' || TRIM(EJECMC_NOMBRE) NOMBCONSEJERO"
   g_str_Parame = g_str_Parame & "   FROM CRE_EJETIP A "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CRE_EJECMC B ON A.EJETIP_CODEJE = B.EJECMC_CODEJE "
   g_str_Parame = g_str_Parame & "  WHERE EJECMC_SITUAC = 1 AND EJETIP_TIPEJE=121 "
   g_str_Parame = g_str_Parame & "  ORDER BY 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Otros, 3) Then
       Exit Sub
   End If
   
   r_rst_Otros.MoveFirst
   Do While Not r_rst_Otros.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT TRIM(A.HIPMAE_CONHIP) AS CONSEJERO,"
      g_str_Parame = g_str_Parame & "        CASE WHEN C.DATGEN_RAZSOC IS NULL THEN 'RECURSOS PROPIOS' ELSE TRIM(C.DATGEN_RAZSOC) END AS PROMOTORA,"
      g_str_Parame = g_str_Parame & "        CASE WHEN B.DATGEN_TITULO IS NULL THEN 'BIEN TERMINADO'   ELSE TRIM(B.DATGEN_TITULO) END AS PROYECTO,"
      g_str_Parame = g_str_Parame & "        COUNT(A.HIPMAE_CODPRD) AS CONTGRUPO"
      g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A"
      g_str_Parame = g_str_Parame & "   LEFT JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.HIPMAE_PRYINM"
      g_str_Parame = g_str_Parame & "   LEFT JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = B.DATGEN_VENTDO AND C.DATGEN_EMPNDO = B.DATGEN_VENNDO"
      g_str_Parame = g_str_Parame & "  WHERE A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "  AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
      g_str_Parame = g_str_Parame & "    AND A.HIPMAE_CONHIP='" & r_rst_Otros!EJETIP_CODEJE & "'"
      g_str_Parame = g_str_Parame & "  GROUP BY A.HIPMAE_CONHIP, C.DATGEN_RAZSOC, B.DATGEN_TITULO"
      g_str_Parame = g_str_Parame & "  ORDER BY CONSEJERO, PROMOTORA, PROYECTO"
      
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Bucle, 3) Then
         Exit Sub
      End If
      
      If (r_rst_Bucle.EOF And r_rst_Bucle.BOF) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont + 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = r_rst_Otros!NOMBCONSEJERO
         
         r_obj_Excel.ActiveSheet.Range("B5:B" & r_int_ConVer).Select
         r_obj_Excel.Selection.Cells.WrapText = True

         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = 0
         
         If s_str_Inicio = "" Then s_str_Inicio = r_int_ConVer
         
         r_obj_Excel.ActiveSheet.Range("G" & s_str_Inicio & ":G" & r_int_ConVer).Select
         r_obj_Excel.ActiveSheet.Range("G" & s_str_Inicio & ":G" & r_int_ConVer).NumberFormat = "#0.#00"
         
         r_int_Cont = r_int_Cont + 1
         r_int_ConVer = r_int_ConVer + 1
                  
         r_obj_Excel.ActiveSheet.Range("A" & r_int_ConVer & ":G" & r_int_ConVer).Interior.Color = RGB(146, 208, 80)
         r_int_ConVer = r_int_ConVer + 1
      End If
      
      r_rst_Otros.MoveNext
   Loop
   
   r_int_Cont1 = 3
   r_int_ConVer1 = 2

   r_obj_Excel.ActiveSheet.Range("A3:A3").Select
   r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
   r_obj_Excel.ActiveSheet.Range("B3:B3").Select
   r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
   r_obj_Excel.ActiveSheet.Range("C3:C3").Select
   r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
   r_obj_Excel.ActiveSheet.Range("D3:D3").Select
   r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
   r_obj_Excel.ActiveSheet.Range("E3:E3").Select
   r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
   r_obj_Excel.ActiveSheet.Range("F3:F3").Select
   r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
   r_obj_Excel.ActiveSheet.Range("G3:G3").Select
   r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
   r_obj_Excel.ActiveSheet.Range("H3:H3").Select
   r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous

   Do While r_int_Cont1 < r_int_ConVer - 1
      If r_int_ConVer1 > 4 Then
         If (r_int_ConVer1 - 2) = 1 Then
            r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 - 1, 1), r_obj_Excel.Cells(r_int_ConVer1 - 1, 7)).Select
            r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
         End If

         r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 - 2, 1), r_obj_Excel.Cells(r_int_ConVer1 - 2, 7)).Select
         r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
         r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 - 1, 1), r_obj_Excel.Cells(r_int_ConVer1 - 1, 7)).Select
         r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
         r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1, 1), r_obj_Excel.Cells(r_int_ConVer1, 7)).Select
         r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      End If
      
      r_obj_Excel.ActiveSheet.Range("A3:A" & 2 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("B3:B" & 2 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("C3:C" & 2 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("D3:D" & 2 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("E3:E" & 2 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("F3:F" & 2 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("G3:G" & 2 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("H3:H" & 2 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      r_int_Cont1 = r_int_Cont1 + 1
      r_int_ConVer1 = r_int_ConVer1 + 1
   Loop
   
   r_int_ConVer1 = r_int_ConVer1 + 2
   r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1, 1), r_obj_Excel.Cells(r_int_ConVer1, 7)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 + 1, 1), r_obj_Excel.Cells(r_int_ConVer1 + 1, 7)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 + 2, 1), r_obj_Excel.Cells(r_int_ConVer1 - 1, 7)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 + 2, 1), r_obj_Excel.Cells(r_int_ConVer1 - 2, 7)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_ConVer1, 4) = "TOTAL DE DESEMBOLSOS : "
      .Cells(r_int_ConVer1, 6) = r_int_Total
      
      .Range("G" & r_int_ConVer1 & ":G" & r_int_ConVer1).Select
      .Range("G" & r_int_ConVer1 & ":G" & r_int_ConVer1).NumberFormat = "#0.#00"
      
      .Cells(r_int_ConVer1, 7) = r_dbl_TotPorc
      .Range(.Cells(r_int_ConVer1, 3), .Cells(r_int_ConVer1, 9)).Font.Bold = True
      .Range(.Cells(r_int_ConVer1, 4), .Cells(r_int_ConVer1, 4)).HorizontalAlignment = xlHAlignRight
      
      .Range("A" & r_int_ConVer1 & ":G" & r_int_ConVer1).Interior.Color = RGB(239, 215, 155)
      .Range("A" & 3 & ":G" & 3).Interior.Color = RGB(213, 239, 245)
      .Range("A1:A2").Select
   End With
            
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
