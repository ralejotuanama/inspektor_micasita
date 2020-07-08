VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_24 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1425
   ClientLeft      =   6930
   ClientTop       =   5010
   ClientWidth     =   5475
   Icon            =   "AteCli_frm_512.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   1440
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   2540
      _StockProps     =   15
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
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
            Height          =   540
            Left            =   630
            TabIndex        =   5
            Top             =   45
            Width           =   4185
            _Version        =   65536
            _ExtentX        =   7382
            _ExtentY        =   952
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes en Trámite en Atención Comercial"
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
            Picture         =   "AteCli_frm_512.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   6
         Top             =   750
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
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
            Picture         =   "AteCli_frm_512.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_512.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4740
            Picture         =   "AteCli_frm_512.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   2
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
   End
End
Attribute VB_Name = "frm_RptSol_24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call fs_GenExc
   
End Sub

Private Sub cmd_Imprim_Click()
   'Confirmación
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Se modifica el puntero para un estado de espera
   Screen.MousePointer = 11
          
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
    
   'Se Muestra las tablas que fueron utilizadas en Crystal Report
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".CRE_PRODUC"
   crp_Imprim.DataFiles(1) = UCase(moddat_g_str_EntDat) & ".CRE_SOLMAE"
   crp_Imprim.DataFiles(2) = UCase(moddat_g_str_EntDat) & ".CLI_DATGEN"
   
   crp_Imprim.SelectionFormula = ""
   
   'Se Filtra por el tipo de producto escogido en el formulario
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_SOLMAE.SOLMAE_CODINS} = 11 AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_SOLMAE.SOLMAE_CONHIP} = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_SOLMAE.SOLMAE_SITUAC} = 1 "
      
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_17.RPT"
   
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
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
End Sub

Private Sub fs_GenExc()

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer

   g_str_Parame = "SELECT * FROM CRE_PRODUC A, CRE_SOLMAE B, CLI_DATGEN C WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = PRODUC_CODIGO AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = DATGEN_TIPDOC AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = DATGEN_NUMDOC AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_SOLMAE.SOLMAE_CONHIP} = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = 11 "
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
      .Cells(1, 3) = "SOLICITUD"
      .Cells(1, 4) = "DOC. IDENTIDAD"
      .Cells(1, 5) = "NOMBRE CLIENTE"
      .Cells(1, 6) = "F. SOLICITUD"
      .Cells(1, 7) = "TIPO EVALUACION"
      .Cells(1, 8) = "CONSEJ. HIPOT."
      .Cells(1, 9) = "V. INMUEBLE S/."
      .Cells(1, 10) = "V. INMUEBLE US$"
      .Cells(1, 11) = "PORC. INICIAL"
      .Cells(1, 12) = "MTO. CREDITO S/."
      .Cells(1, 13) = "MTO. CREDITO US$"
            
      .Range(.Cells(1, 1), .Cells(1, 13)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 13)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      
      .Columns("B").ColumnWidth = 30
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 40
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 40
      
      .Columns("H").ColumnWidth = 15
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 15
      .Columns("J").ColumnWidth = 15
      .Columns("K").ColumnWidth = 15
      .Columns("L").ColumnWidth = 15
      .Columns("M").ColumnWidth = 15
   End With
   
   g_rst_Princi.MoveFirst
     
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
      End If
           
      If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Or g_rst_Princi!SOLMAE_COMVTA_DOL > 0 Then
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_SOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00") * 100) + "%"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_DOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00") * 100) + "%"
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = CStr(0) + "%"
      End If
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
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

