VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_RptSol_30 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   8010
   ClientTop       =   4710
   ClientWidth     =   5430
   Icon            =   "AteCli_frm_536.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2265
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5445
      _Version        =   65536
      _ExtentX        =   9604
      _ExtentY        =   3995
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
         TabIndex        =   5
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
            Height          =   525
            Left            =   660
            TabIndex        =   6
            Top             =   30
            Width           =   4635
            _Version        =   65536
            _ExtentX        =   8176
            _ExtentY        =   926
            _StockProps     =   15
            Caption         =   "Resumen de Solicitudes Desembolsadas por Consejero Hipotecario"
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
            Picture         =   "AteCli_frm_536.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   7
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
            Picture         =   "AteCli_frm_536.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_536.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   8
         Top             =   1440
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
         _ExtentY        =   1349
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
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1320
            TabIndex        =   0
            Top             =   60
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
            Left            =   1320
            TabIndex        =   1
            Top             =   390
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
            TabIndex        =   10
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   225
            Left            =   60
            TabIndex        =   9
            Top             =   390
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

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
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

'Private Sub cmd_Imprim_Click()
'   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
'      MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_FecIni)
'      Exit Sub
'   End If
'
'    'Confirmación
'   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
'      Exit Sub
'   End If
'
'   Screen.MousePointer = 11
'   Call fs_GenTmp
'   Screen.MousePointer = 0
'
'   'Se envia la cadena de conexión
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'   crp_Imprim.DataFiles(0) = "RPT_ESTCON"
'   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
'   crp_Imprim.DataFiles(2) = "CRE_EJECMC"
'
'   'Se selecciona la formula
'   crp_Imprim.SelectionFormula = "{RPT_ESTCON.ESTCON_TERCRE} = '" & modgen_g_str_NombPC & "' "
'
'   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
'   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RESDES_01.RPT"
'   crp_Imprim.Action = 1
'End Sub

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
End Sub

Private Sub fs_GenTmp()
Dim r_int_TotOpe     As Integer
Dim r_dbl_MtoSol     As Double
Dim r_dbl_MtoDol     As Double
Dim r_str_ConHip     As String
Dim r_str_CodPrd     As String

   'Borrando Temporal
   g_str_Parame = "DELETE FROM RPT_ESTCON WHERE "
   g_str_Parame = g_str_Parame & "ESTCON_TERCRE = '" & modgen_g_str_NombPC & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   'Para determinar los
   g_str_Parame = "SELECT HIPMAE_CONHIP, HIPMAE_CODPRD, HIPMAE_MONEDA, COUNT(HIPMAE_NUMOPE) AS TOTOPE, SUM(HIPMAE_MTOPRE) AS TOTPRE FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CONHIP, HIPMAE_CODPRD, HIPMAE_MONEDA "
   g_str_Parame = g_str_Parame & "ORDER BY HIPMAE_CONHIP ASC, HIPMAE_CODPRD ASC, HIPMAE_MONEDA ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_ConHip = Trim(g_rst_Princi!HIPMAE_CONHIP)
         r_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
         r_int_TotOpe = 0
         r_dbl_MtoSol = 0
         r_dbl_MtoDol = 0
      
         Do While Not g_rst_Princi.EOF And r_str_ConHip = Trim(g_rst_Princi!HIPMAE_CONHIP) And r_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
            r_int_TotOpe = r_int_TotOpe + g_rst_Princi!TOTOPE
            If g_rst_Princi!HIPMAE_MONEDA = 1 Then
               r_dbl_MtoSol = r_dbl_MtoSol + g_rst_Princi!TOTPRE
            Else
               r_dbl_MtoDol = r_dbl_MtoDol + g_rst_Princi!TOTPRE
            End If
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      
         'Insertar Línea por Consejero y Producto en Tabla Temporal
         g_str_Parame = "INSERT INTO RPT_ESTCON ("
         g_str_Parame = g_str_Parame & "ESTCON_TERCRE, "
         g_str_Parame = g_str_Parame & "ESTCON_CONHIP, "
         g_str_Parame = g_str_Parame & "ESTCON_CODPRD, "
         g_str_Parame = g_str_Parame & "ESTCON_TOTOPE, "
         g_str_Parame = g_str_Parame & "ESTCON_MTOSOL, "
         g_str_Parame = g_str_Parame & "ESTCON_MTODOL, "
         g_str_Parame = g_str_Parame & "ESTCON_FECINI, "
         g_str_Parame = g_str_Parame & "ESTCON_FECFIN) "
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_ConHip & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_CodPrd & "', "
         g_str_Parame = g_str_Parame & CStr(r_int_TotOpe) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_MtoSol) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_MtoDol) & ", "
         g_str_Parame = g_str_Parame & "'" & ipp_FecIni.Text & "', "
         g_str_Parame = g_str_Parame & "'" & ipp_FecFin.Text & "') "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
             Exit Sub
         End If
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "AÑO DESEMBOLSO"
      .Cells(1, 3) = "MES DESEMBOLSO"
      .Cells(1, 4) = "FECHA DESEMBOLSO"
      .Cells(1, 5) = "CONSEJERO"
      .Cells(1, 6) = "OPERACION"
      .Cells(1, 7) = "PRODUCTO"
      .Cells(1, 8) = "MONEDA"
      .Cells(1, 9) = "MONTO"
      .Cells(1, 10) = "MODALIDAD"
      .Cells(1, 11) = "PROYECTO"
      .Cells(1, 12) = "PROMOTOR"
      .Cells(1, 13) = "VINCULADO"
      
      .Range(.Cells(1, 1), .Cells(1, 13)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 13)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 6
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 17
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 17
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 19
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 16
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 14
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 14
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 14
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 12
      .Columns("I").HorizontalAlignment = xlHAlignRight
      
      .Columns("J").ColumnWidth = 25
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 44
      '.Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 50
      '.Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 12
      .Columns("M").HorizontalAlignment = xlHAlignCenter
   End With
     
   r_int_ConVer = 2
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUBSTR(A.HIPMAE_FECDES,1,4)            AS ANIO_DESEMBOLSO, "
   g_str_Parame = g_str_Parame & "       SUBSTR(A.HIPMAE_FECDES,5,2)            AS MES_DESEMBOLSO, "
   g_str_Parame = g_str_Parame & "       HIPMAE_FECDES                          AS FECHA_DESEMBOLSO, "
   g_str_Parame = g_str_Parame & "       TRIM(A.HIPMAE_CONHIP)                  AS CONSEJERO, "
   g_str_Parame = g_str_Parame & "       HIPMAE_NUMOPE                          AS OPERACION, "
   g_str_Parame = g_str_Parame & "       CASE WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrCRC & ")  THEN   'CRC-PBP' "
   g_str_Parame = g_str_Parame & "            WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrCME & ")  THEN   'CME' "
   g_str_Parame = g_str_Parame & "            WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN   'MICASITA' "
   g_str_Parame = g_str_Parame & "            WHEN A.HIPMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN   'MIVIVIENDA' "
   g_str_Parame = g_str_Parame & "       END                                    AS PRODUCTO, "
   g_str_Parame = g_str_Parame & "       DECODE(HIPMAE_MONEDA, 1, 'SOLES', 'DOLARES AMERICANOS')                 AS TIPO_MONEDA, "
   g_str_Parame = g_str_Parame & "       HIPMAE_TOTPRE                          AS MONTO_PRESTAMO, "
   g_str_Parame = g_str_Parame & "       HIPMAE_CODMOD                          AS MODALIDAD, "
   g_str_Parame = g_str_Parame & "       HIPMAE_CODPRD                          , "
   g_str_Parame = g_str_Parame & "       DECODE(NVL(TRIM(B.DATGEN_TITULO), ''), '', 'SIN DATOS', TRIM(B.DATGEN_TITULO)) AS PROYECTO, "
   g_str_Parame = g_str_Parame & "       DECODE(NVL(TRIM(B.DATGEN_TITULO), ''), '', TRIM(D.SOLINM_RAZSOC_PRO) , TRIM(C.DATGEN_RAZSOC)) AS PROMOTOR, "
   g_str_Parame = g_str_Parame & "       DECODE(A.HIPMAE_PRYMCS, 1, 'SI', 'NO') AS VINCULADO "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.HIPMAE_PRYINM "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = B.DATGEN_VENTDO AND C.DATGEN_EMPNDO = B.DATGEN_VENNDO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLINM D ON D.SOLINM_NUMSOL = A.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE A.HIPMAE_SITUAC IN (2, 6, 9) "
   g_str_Parame = g_str_Parame & "   AND A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & " ORDER BY ANIO_DESEMBOLSO, MES_DESEMBOLSO, HIPMAE_FECDES "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Insertar Línea por Consejero y Producto en Tabla Temporal
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = g_rst_Princi!ANIO_DESEMBOLSO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = g_rst_Princi!MES_DESEMBOLSO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = g_rst_Princi!FECHA_DESEMBOLSO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = g_rst_Princi!CONSEJERO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = g_rst_Princi!OPERACION
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = g_rst_Princi!PRODUCTO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = g_rst_Princi!TIPO_MONEDA
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!MONTO_PRESTAMO, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!HIPMAE_CODPRD), Trim(g_rst_Princi!MODALIDAD)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = g_rst_Princi!PROYECTO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = g_rst_Princi!PROMOTOR
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = g_rst_Princi!VINCULADO
         
         r_int_ConVer = r_int_ConVer + 1
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
