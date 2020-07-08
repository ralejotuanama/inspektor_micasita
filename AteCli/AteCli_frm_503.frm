VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1440
   ClientLeft      =   7575
   ClientTop       =   6225
   ClientWidth     =   5055
   Icon            =   "AteCli_frm_503.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5085
      _Version        =   65536
      _ExtentX        =   8969
      _ExtentY        =   2566
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
         TabIndex        =   4
         Top             =   30
         Width           =   4995
         _Version        =   65536
         _ExtentX        =   8811
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
            Height          =   495
            Left            =   570
            TabIndex        =   5
            Top             =   60
            Width           =   3765
            _Version        =   65536
            _ExtentX        =   6641
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes en Trámite con Aprobación Condicionada"
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
            Picture         =   "AteCli_frm_503.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   6
         Top             =   750
         Width           =   4995
         _Version        =   65536
         _ExtentX        =   8811
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
            Picture         =   "AteCli_frm_503.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_503.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4380
            Picture         =   "AteCli_frm_503.frx":0A62
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
Attribute VB_Name = "frm_RptSol_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub cmd_ExpExc_Click()
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call fs_GenExc
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_str_InsApr_Con     As String
   Dim r_str_SitSol         As String
   Dim r_str_InsSol         As String
   Dim r_str_SitCon         As String
   
   'Confirmacion
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Se cambia el estado del mouse en espera
   Screen.MousePointer = 11
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat

   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "TRA_SEGCON"
   crp_Imprim.DataFiles(1) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"

   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_APRCON_01.RPT"
   
   crp_Imprim.SelectionFormula = "{TRA_SEGCON.SEGCON_SITUAC} = 1 AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_SOLMAE.SOLMAE_CONHIP} = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "({CRE_SOLMAE.SOLMAE_SITUAC} = 1 OR {CRE_SOLMAE.SOLMAE_SITUAC} = 2)"
   
   'Se escoge donde se destinara el reporte
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
      
   'Se cambia el estado del puntero a estado normal
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

'   g_str_Parame = "SELECT SOLMAE_CODPRD, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, A.SEGFECCRE, SOLMAE_CODINS, SOLMAE_FECSOL, SEGCON_CODINS, SOLMAE_CONHIP, SEGCON_OBSCON FROM TRA_SEGCON A, CRE_SOLMAE B, CLI_DATGEN C WHERE "
'   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SEGCON_NUMSOL AND "
'   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = DATGEN_TIPDOC AND "
'   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = DATGEN_NUMDOC AND "

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PRODUC_DESCRI, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_TIPMON, "
   g_str_Parame = g_str_Parame & "        (TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ' ' || TRIM(DATGEN_NOMBRE)) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "        F.PARDES_DESCRI AS INSTANCIA, SOLMAE_CONHIP , SOLMAE_FECSOL, "
   g_str_Parame = g_str_Parame & "        D.SEGFECCRE, D.SEGCON_CODINS, D.SEGCON_OBSCON, E.PARDES_DESCRI AS INSTANCIA_ACTUAL, F.PARDES_DESCRI AS INSTANCIA_APRCON "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_SEGCON D ON D.SEGCON_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES E ON E.PARDES_CODITE = A.SOLMAE_CODINS AND E.PARDES_CODGRP = '002'"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES F ON F.PARDES_CODITE = D.SEGCON_CODINS AND F.PARDES_CODGRP = '002'"

   g_str_Parame = g_str_Parame & "  WHERE "

   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "(SOLMAE_SITUAC = 1 OR SOLMAE_SITUAC = 2) AND "
   g_str_Parame = g_str_Parame & "SEGCON_SITUAC = 1 "
            
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CODPRD ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
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
      .Cells(1, 4) = "NOMBRE CLIENTE"
      .Cells(1, 5) = "F. SOLICITUD"
      .Cells(1, 6) = "INSTANCIA ACTUAL"
      .Cells(1, 7) = "F. APROB. CONDIC."
      .Cells(1, 8) = "INSTANCIA APROB. CONDIC."
      .Cells(1, 9) = "CONSEJ. HIPOT."
      .Cells(1, 10) = "OBSERVACION"
   
      .Range(.Cells(1, 1), .Cells(1, 11)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 11)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 32
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 42
            
      .Columns("E").ColumnWidth = 24
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 38
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 24
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 46
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 14
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 80
   End With
   
   g_rst_Princi.MoveFirst
     
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
               
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI) 'moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!CLIENTE) 'moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!INSTANCIA_ACTUAL) 'moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SOLMAE_CODINS))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!INSTANCIA_APRCON) 'moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGCON_CODINS))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!SEGCON_OBSCON)
                                
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

