VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_19 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   6315
   ClientTop       =   5415
   ClientWidth     =   6930
   Icon            =   "AteCli_frm_518.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2265
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
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
         TabIndex        =   6
         Top             =   30
         Width           =   6855
         _Version        =   65536
         _ExtentX        =   12091
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
            TabIndex        =   7
            Top             =   30
            Width           =   6105
            _Version        =   65536
            _ExtentX        =   10769
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes en Trámite con Observaciones Vigentes"
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
            Height          =   285
            Left            =   660
            TabIndex        =   11
            Top             =   300
            Width           =   6105
            _Version        =   65536
            _ExtentX        =   10769
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Por Consejero Hipotecario"
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
            Picture         =   "AteCli_frm_518.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   750
         Width           =   6855
         _Version        =   65536
         _ExtentX        =   12091
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
            Left            =   6240
            Picture         =   "AteCli_frm_518.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_518.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_518.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
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
         Height          =   735
         Left            =   30
         TabIndex        =   9
         Top             =   1440
         Width           =   6855
         _Version        =   65536
         _ExtentX        =   12091
         _ExtentY        =   1296
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
            Left            =   1920
            TabIndex        =   1
            Top             =   390
            Width           =   2685
         End
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   4875
         End
         Begin VB.Label Label4 
            Caption         =   "Consejero Hipotecario:"
            Height          =   255
            Left            =   60
            TabIndex        =   10
            Top             =   90
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_ConHip()   As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub chk_ConHip_Click()
   If chk_ConHip.Value = 1 Then
      cmb_ConHip.ListIndex = -1
      cmb_ConHip.Enabled = False
      Call gs_SetFocus(cmd_Imprim)
   ElseIf chk_ConHip.Value = 0 Then
      cmb_ConHip.Enabled = True
      Call gs_SetFocus(cmb_ConHip)
   End If
End Sub

Private Sub cmb_ConHip_Click()
   Call gs_SetFocus(cmd_Imprim)
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
       
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
Dim r_dbl_FecObs     As Double
Dim r_str_DesIns     As String
         
   'Validación
   If chk_ConHip = 0 Then
      If cmb_ConHip.ListIndex = -1 Then
         MsgBox "Debe seleccionar a un Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ConHip)
         Exit Sub
      End If
   End If
      
   'Proceso
   Screen.MousePointer = 11
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SEGOBS "
   g_str_Parame = g_str_Parame & " WHERE SEGOBS_NOMRPT = 'ATE_RPTSOL_12.RPT' "
   g_str_Parame = g_str_Parame & "   AND SEGOBS_TERCRE = '" & modgen_g_str_NombPC & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(D.PRODUC_DESCRI) AS PRODUCTO, SOLMAE_NUMERO, C.SEGDET_CODINS, TRIM(E.PARDES_DESCRI) AS INSTANCIA, "
   g_str_Parame = g_str_Parame & "       TRIM(B.DATGEN_APEPAT)||' '||TRIM(B.DATGEN_APEMAT)||' '||TRIM(B.DATGEN_NOMBRE) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "       SOLMAE_FECSOL, SEGDET_FECOCU, SOLMAE_CONHIP, SOLMAE_TIPMON, SOLMAE_COMVTA_SOL, SOLMAE_COMVTA_DOL, "
   g_str_Parame = g_str_Parame & "       SOLMAE_APOPRO_SOL , SOLMAE_APOPRO_DOL, SOLMAE_MTOPRE_MPR, SEGDET_OBSERV, SOLMAE_TITTDO, SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.SOLMAE_TITTDO AND B.DATGEN_NUMDOC = A.SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & " INNER JOIN TRA_SEGDET C ON C.SEGDET_NUMSOL = A.SOLMAE_NUMERO AND C.SEGDET_CODOCU = 21 AND C.SEGFECACT = 0 "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC D ON D.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = '002' AND E.PARDES_CODITE = C.SEGDET_CODINS "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_SITUAC = 1 "
   If chk_ConHip.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND SOLMAE_CODPRD = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "' "
   End If
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "   AND SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' "
   End If
   g_str_Parame = g_str_Parame & " ORDER BY SOLMAE_CONHIP ASC, CLIENTE ASC, SEGDET_CODINS "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
            
         r_dbl_FecObs = ff_FecObs(g_rst_Princi!SOLMAE_NUMERO)
         r_str_DesIns = Trim(g_rst_Princi!INSTANCIA)
         
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_SEGOBS("
         g_str_Parame = g_str_Parame & "SEGOBS_NOMRPT, "
         g_str_Parame = g_str_Parame & "SEGOBS_FECCRE, "
         g_str_Parame = g_str_Parame & "SEGOBS_HORCRE, "
         g_str_Parame = g_str_Parame & "SEGOBS_TERCRE, "
         g_str_Parame = g_str_Parame & "SEGOBS_NUMSOL, "
         g_str_Parame = g_str_Parame & "SEGOBS_FECOBS, "
         g_str_Parame = g_str_Parame & "SEGOBS_DESINS, "
         g_str_Parame = g_str_Parame & "SEGOBS_OBSERV) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_12.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & CStr(r_dbl_FecObs) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_DesIns & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SEGDET_OBSERV & "' ) "
         
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
      
   Screen.MousePointer = 0
   
   'Confirmación
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
     
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = UCase(moddat_g_str_EntDat) & ".CRE_PRODUC"
   crp_Imprim.DataFiles(2) = UCase(moddat_g_str_EntDat) & ".CLI_DATGEN"
   crp_Imprim.DataFiles(3) = UCase(moddat_g_str_EntDat) & ".RPT_SEGOBS"
   
   crp_Imprim.SelectionFormula = "{RPT_SEGOBS.SEGOBS_NOMRPT} = 'ATE_RPTSOL_12.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SEGOBS.SEGOBS_TERCRE} = '" & modgen_g_str_NombPC & "' "
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_12.RPT"
   crp_Imprim.Action = 1
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

Private Function ff_FecObs(ByVal p_NumSol As String) As Double
   ff_FecObs = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_SEGDET "
   g_str_Parame = g_str_Parame & " WHERE SEGDET_NUMSOL = '" & p_NumSol & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         ff_FecObs = g_rst_Listas!SEGDET_FECOCU
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(D.PRODUC_DESCRI) AS PRODUCTO, SOLMAE_NUMERO, C.SEGDET_CODINS, TRIM(E.PARDES_DESCRI) AS INSTANCIA, "
   g_str_Parame = g_str_Parame & "       TRIM(B.DATGEN_APEPAT)||' '||TRIM(B.DATGEN_APEMAT)||' '||TRIM(B.DATGEN_NOMBRE) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "       SOLMAE_FECSOL, SEGDET_FECOCU, SOLMAE_CONHIP, SOLMAE_TIPMON, SOLMAE_COMVTA_SOL, SOLMAE_COMVTA_DOL, "
   g_str_Parame = g_str_Parame & "       SOLMAE_APOPRO_SOL , SOLMAE_APOPRO_DOL, SOLMAE_MTOPRE_MPR, SEGDET_OBSERV, SOLMAE_TITTDO, SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.SOLMAE_TITTDO AND B.DATGEN_NUMDOC = A.SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & " INNER JOIN TRA_SEGDET C ON C.SEGDET_NUMSOL = A.SOLMAE_NUMERO AND C.SEGDET_CODOCU = 21 AND C.SEGFECACT = 0 "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC D ON D.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = '002' AND E.PARDES_CODITE = C.SEGDET_CODINS "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_SITUAC = 1 "
   If chk_ConHip.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND SOLMAE_CODPRD = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "' "
   End If
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "   AND SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' "
   End If
   g_str_Parame = g_str_Parame & " ORDER BY SOLMAE_CONHIP ASC, CLIENTE ASC, SEGDET_CODINS "
   
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
      .Cells(1, 7) = "F. SOLICITUD"
      .Cells(1, 8) = "F. OBSERVACION"
      .Cells(1, 9) = "TIP. DE MONEDA"
      .Cells(1, 10) = "V. INMUEBLE S/."
      .Cells(1, 11) = "V. INMUEBLE US$."
      .Cells(1, 12) = "PORC. INICIAL"
      .Cells(1, 13) = "MTO. CREDITO S/."
      .Cells(1, 14) = "MTO. CREDITO US$."
      .Cells(1, 15) = "INSTANCIA ACTUAL"
      .Cells(1, 16) = "OBSERVACION"
   
      .Range(.Cells(1, 1), .Cells(1, 16)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 16)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 20
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 32
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 40
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 15
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 22
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 16
      .Columns("K").ColumnWidth = 16
      .Columns("L").ColumnWidth = 13
      .Columns("M").ColumnWidth = 17
      .Columns("N").ColumnWidth = 17
      .Columns("O").ColumnWidth = 52
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 200
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!PRODUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!CLIENTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGDET_FECOCU)))
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = "NUEVOS SOLES"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = "DOLARES AMERICANOS"
      End If
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = 0
      End If
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = 0
      End If
      If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Or g_rst_Princi!SOLMAE_COMVTA_DOL > 0 Then
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_SOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00") * 100) + "%"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_DOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00") * 100) + "%"
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CStr(0) + "%"
      End If
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = 0
      End If
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = 0
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Trim(g_rst_Princi!INSTANCIA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Trim(g_rst_Princi!SEGDET_OBSERV)
      
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

