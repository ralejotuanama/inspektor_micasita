VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_27 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   6870
   ClientTop       =   6570
   ClientWidth     =   5460
   Icon            =   "AteCli_frm_522.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2235
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5505
      _Version        =   65536
      _ExtentX        =   9710
      _ExtentY        =   3942
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
            Height          =   255
            Left            =   630
            TabIndex        =   7
            Top             =   30
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes en Evaluación Crediticia"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
            Height          =   270
            Left            =   630
            TabIndex        =   11
            Top             =   300
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Por Consejero Hipotecario"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
            Picture         =   "AteCli_frm_522.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
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
            Picture         =   "AteCli_frm_522.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_522.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4770
            Picture         =   "AteCli_frm_522.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
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
         Height          =   705
         Left            =   30
         TabIndex        =   9
         Top             =   1440
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   1244
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
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3375
         End
         Begin VB.CheckBox chk_ConHip 
            Caption         =   "Todos los Consejeros Hipotecarios"
            Height          =   285
            Left            =   1920
            TabIndex        =   1
            Top             =   390
            Width           =   2835
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   60
            TabIndex        =   10
            Top             =   90
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
         MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ConHip)
         Exit Sub
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
   Dim r_dbl_FecIng     As Double

   'Validación
   If chk_ConHip = 0 Then
      If cmb_ConHip.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ConHip)
         Exit Sub
      End If
   End If
   
   'Confirmación
   If MsgBox("¿Está seguro de imprimir el Reporte de Solicitudes en Evaluación Crediticia?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(Date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   'Se elimina los datos de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
     
   'Leyendo Tabla de solicitudes
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   'Si no escogio todos los Productos
   If chk_ConHip.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "' AND "
   End If
        
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = 21 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
   
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
      
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_DesOcu = moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
                 
         'Para obtener la Fecha de Ingreso a Aprobación Crediticia
         r_dbl_FecIng = ff_FecIng(g_rst_Princi!SOLMAE_NUMERO)
      
         If r_dbl_FecIng <> 0 Then
            'Insertando Registro
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "INSERT INTO RPT_SOLTRA("
            g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT, "
            g_str_Parame = g_str_Parame & "SOLTRA_FECCRE, "
            g_str_Parame = g_str_Parame & "SOLTRA_HORCRE, "
            g_str_Parame = g_str_Parame & "SOLTRA_TERCRE, "
            g_str_Parame = g_str_Parame & "SOLTRA_NUMSOL, "
            g_str_Parame = g_str_Parame & "SOLTRA_CODOCU, "
            g_str_Parame = g_str_Parame & "SOLTRA_FECING) "
            
            g_str_Parame = g_str_Parame & "VALUES ("
            g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_20.RPT" & "', "
            g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
            g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
            g_str_Parame = g_str_Parame & "'" & r_str_DesOcu & "', "
            g_str_Parame = g_str_Parame & CStr(r_dbl_FecIng) & ") "
                     
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               Exit Sub
            End If
         End If
      
         g_rst_Princi.MoveNext
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
   Else
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se Muestra las tablas que fueron utilizadas en Crystal Report
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = UCase(moddat_g_str_EntDat) & ".CRE_PRODUC"
   crp_Imprim.DataFiles(2) = UCase(moddat_g_str_EntDat) & ".CLI_DATGEN"
   crp_Imprim.DataFiles(3) = UCase(moddat_g_str_EntDat) & ".RPT_SOLTRA"
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_20.RPT"
   
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
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
       
End Sub
Private Function ff_FecIng(ByVal p_Numsol As String) As Double
   
   ff_FecIng = 0
      
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & p_Numsol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
         ff_FecIng = g_rst_Listas!SEGFECCRE
                  
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
   
End Function

Private Sub fs_GenExc()

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer

   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, CLI_DATGEN B, RPT_SOLTRA C WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = DATGEN_TIPDOC AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = DATGEN_NUMDOC AND "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLTRA_NUMSOL AND "
   
   If chk_ConHip.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "' AND "
   End If
      
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = 21 AND "
   g_str_Parame = g_str_Parame & "SOLTRA_FECING <> 0 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CONHIP ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC "
   
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
      .Cells(1, 2) = "CONSEJ. HIPOT"
      .Cells(1, 3) = "PRODUCTO"
      .Cells(1, 4) = "SOLICITUD"
      .Cells(1, 5) = "DOC. IDENTIDAD"
      .Cells(1, 6) = "NOMBRE CLIENTE"
      .Cells(1, 7) = "F. SOLICITUD"
      .Cells(1, 8) = "F.INGR. EV. CRED."
      .Cells(1, 9) = "SITUACION EN INSTANCIA"
      .Cells(1, 10) = "TIPO EVALUACION"
      .Cells(1, 11) = "TIP. MONEDA"
      .Cells(1, 12) = "V. INMUEBLE S/."
      .Cells(1, 13) = "V. INMUEBLE US$."
      .Cells(1, 14) = "PORC. INICIAL"
      .Cells(1, 15) = "MTO. CREDITO S/."
      .Cells(1, 16) = "MTO. CREDITO US$."
            
      .Range(.Cells(1, 1), .Cells(1, 16)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 16)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      
      .Columns("B").ColumnWidth = 14
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 30
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 16
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 45
      
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 17
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 33
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 28
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 17
      .Columns("L").ColumnWidth = 17
      .Columns("M").ColumnWidth = 13
      .Columns("N").ColumnWidth = 17
      .Columns("O").ColumnWidth = 18
      .Columns("P").ColumnWidth = 18
            
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
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(ff_FecIng(g_rst_Princi!SOLMAE_NUMERO)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
                      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = "NUEVOS SOLES"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = "DOLARES AMERICANOS"
      End If
                      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = 0
      End If
                  
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = 0
      End If
      
      If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Or g_rst_Princi!SOLMAE_COMVTA_DOL > 0 Then
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_SOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00") * 100) + "%"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_DOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00") * 100) + "%"
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = CStr(0) + "%"
      End If
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
      End If
                  
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
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



