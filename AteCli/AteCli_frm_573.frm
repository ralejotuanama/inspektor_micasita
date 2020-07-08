VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_47 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6240
   Icon            =   "AteCli_frm_573.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   2145
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   3784
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
         Width           =   6165
         _Version        =   65536
         _ExtentX        =   10874
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
         Begin Threed.SSPanel ssp_TipCon 
            Height          =   555
            Left            =   630
            TabIndex        =   6
            Top             =   30
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes"
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
            Picture         =   "AteCli_frm_573.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   7
         Top             =   750
         Width           =   6165
         _Version        =   65536
         _ExtentX        =   10874
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
            Picture         =   "AteCli_frm_573.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_573.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5550
            Picture         =   "AteCli_frm_573.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   3
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   675
         Left            =   30
         TabIndex        =   8
         Top             =   1420
         Width           =   6165
         _Version        =   65536
         _ExtentX        =   10874
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
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   4635
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Reporte:"
            Height          =   255
            Left            =   60
            TabIndex        =   9
            Top             =   210
            Width           =   1290
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_47"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub cmd_ExpExc_Click()
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      Call fs_GenExc_AteCli
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      Call fs_GenExc_EvaCre
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
      Call fs_GenExc_AceCre
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Then
      Call fs_GenExc_TasSeg
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 5 Then
      Call fs_GenExc_EvaLeg
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 6 Then
      Call fs_GenExc_PMVCof
   End If
End Sub

Private Sub cmb_TipRep_Click()
   If cmb_TipRep.ListIndex <> -1 Then
      If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
         ssp_TipCon.Caption = "Reporte de Solicitudes en Trámite en Atención Comercial"
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
         ssp_TipCon.Caption = "Reporte de Solicitudes en Evaluación Crediticia"
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
         ssp_TipCon.Caption = "Reporte de Solicitudes en Aceptación Crediticia y Trámites de Cliente"
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Then
         ssp_TipCon.Caption = "Reporte de Solicitudes en Tasación y Seguros"
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 5 Then
         ssp_TipCon.Caption = "Reporte de Solicitudes en Evaluación Legal"
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 6 Then
         ssp_TipCon.Caption = "Reporte de Solicitudes en Pólizas y Mivivienda-Cofide"
      End If
   End If
   Call gs_SetFocus(cmd_Imprim)
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipRep_Click
   End If
End Sub

Private Sub cmd_Imprim_Click()
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   'Confirmación
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Se modifica el puntero para un estado de espera
   Screen.MousePointer = 11
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      Call fs_GenImp_AteCli
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      Call fs_GenImp_EvaCre
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
      Call fs_GenImp_AceCre
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Then
      Call fs_GenImp_TasSeg
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 5 Then
      Call fs_GenImp_EvaLeg
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 6 Then
      Call fs_GenImp_PMVCof
   End If
     
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   
   cmb_TipRep.AddItem "Solicitudes en Trámite en Atención Comercial"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
   
   cmb_TipRep.AddItem "Solicitudes en Evaluación Crediticia"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2
   
   cmb_TipRep.AddItem "Solicitudes en Aceptación Crediticia y Trámites de Cliente"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 3
   
   cmb_TipRep.AddItem "Solicitudes en Tasación y Seguros"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 4
   
   cmb_TipRep.AddItem "Solicitudes en Evaluación Legal"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 5
   
   cmb_TipRep.AddItem "Solicitudes en Pólizas y Mivivienda-Cofide"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 6
   
   cmb_TipRep.ListIndex = -1
      
   Call gs_CentraForm(Me)
End Sub

Private Function ff_IngIns(ByVal p_NumSol As String, ByVal p_CodIns As Integer) As Double
   ff_IngIns = 0
      
   g_str_Parame = "SELECT SEGUIM_FECINI FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(p_CodIns) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      ff_IngIns = g_rst_Listas!SEGUIM_FECINI
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function ff_SitIns(ByVal p_NumSol As String, ByVal p_CodIns As Integer) As String
   Dim r_rst_Genera     As ADODB.Recordset
   
   ff_SitIns = ""
      
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(p_CodIns) & " "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      ff_SitIns = moddat_gf_Consulta_ParDes("023", CStr(r_rst_Genera!SEGUIM_SITUAC))
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Private Sub fs_GenImp_AteCli()
   'Se envia la cadena de conexión
   crp_Imprim.Reset
   crp_Imprim.WindowState = crptMaximized
   crp_Imprim.WindowTitle = "Presentación Preliminar"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.WindowShowRefreshBtn = True
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se Muestra las tablas que fueron utilizadas en Crystal Report
   crp_Imprim.DataFiles(0) = "CRE_PRODUC"
   crp_Imprim.DataFiles(1) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   
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
End Sub

Private Sub fs_GenImp_EvaCre()
   Dim r_str_DesOcu     As String
   Dim r_dbl_FecIng     As Double
          
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
   
   'Se elimina los datos de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT = 'ATE_RPTSOL_19.RPT' AND "
   g_str_Parame = g_str_Parame & "SOLTRA_TERCRE = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
      
   'Leyendo Tabla de solicitudes
   'g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_NUMERO, B.PARDES_DESCRI AS SITUACION "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES B ON B.PARDES_CODITE = A.SOLMAE_SITINS AND B.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "  WHERE SOLMAE_CODINS = 21 AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
      
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_DesOcu = Trim(g_rst_Princi!SITUACION) 'moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
                  
         'Para obtener la Fecha de Ingreso a Aprobación Crediticia
         r_dbl_FecIng = ff_FecIng(g_rst_Princi!SOLMAE_NUMERO)
         
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
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_19.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_DesOcu & "', "
         g_str_Parame = g_str_Parame & CStr(r_dbl_FecIng) & ") "
                  
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
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
            
   crp_Imprim.Reset
   crp_Imprim.WindowState = crptMaximized
   crp_Imprim.WindowTitle = "Presentación Preliminar"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.WindowShowRefreshBtn = True
   
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se Muestra las tablas que fueron utilizadas en Crystal Report
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_SOLTRA"
   crp_Imprim.DataFiles(4) = "TRA_SEGUIM"
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_19.RPT"
   
   crp_Imprim.SelectionFormula = "{RPT_SOLTRA.SOLTRA_NOMRPT} = 'ATE_RPTSOL_19.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SOLTRA.SOLTRA_TERCRE} = '" & modgen_g_str_NombPC & "' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{TRA_SEGUIM.SEGUIM_CODINS} = 21 "
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Function ff_FecIng(ByVal p_NumSol As String) As Double
   ff_FecIng = 0
      
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 21 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      ff_FecIng = g_rst_Listas!SEGUIM_FECINI
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_GenImp_AceCre()
   Dim r_str_DesOcu     As String
   Dim r_dbl_IngAce     As Double
   Dim r_dbl_IngTra     As Double
   Dim r_dbl_ImpGas     As Double
   Dim r_dbl_FecPag     As Double
           
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
   
   'Se elimina los datos de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT = 'ATE_RPTSOL_22.RPT' AND "
   g_str_Parame = g_str_Parame & "SOLTRA_TERCRE = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
      
   'Leyendo Tabla de solicitudes
   'g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_NUMERO, B.PARDES_DESCRI AS SITUACION, C.SEGUIM_FECINI AS FECING_ACECRE, "
   g_str_Parame = g_str_Parame & "        D.SEGUIM_FECINI AS FECING_TRACLI, E.GASADM_PAGFEC "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES B ON B.PARDES_CODITE = A.SOLMAE_SITINS AND B.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_SEGUIM C ON C.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND C.SEGUIM_CODINS = '31'"
   g_str_Parame = g_str_Parame & "         LEFT JOIN TRA_SEGUIM D ON D.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND D.SEGUIM_CODINS = '32'"
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT GASADM_NUMSOL, SUM(GASADM_PAGIMP) AS PAGIMP, GASADM_PAGFEC FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "                           GASADM_SITUAC = 1 GROUP BY GASADM_NUMSOL, GASADM_PAGFEC) E ON E.GASADM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & " WHERE (SOLMAE_CODINS = 31 OR SOLMAE_CODINS = 32) AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & " SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & " SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
      
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_DesOcu = Trim(g_rst_Princi!SITUACION) 'moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
         
         'Para obtener Fecha de Ingreso a Aceptación Crediticia
         r_dbl_IngAce = IIf(IsNull(g_rst_Princi!FECING_ACECRE), 0, g_rst_Princi!FECING_ACECRE) 'ff_IngIns(g_rst_Princi!SOLMAE_NUMERO, 31)
         
         'Para obtener Fecha de Ingreso a Trámites de Cliente
         r_dbl_IngTra = IIf(IsNull(g_rst_Princi!FECING_TRACLI), 0, g_rst_Princi!FECING_TRACLI) 'ff_IngIns(g_rst_Princi!SOLMAE_NUMERO, 32)
         
         'Para obtener Fecha de Pago de Gastos de Cierre
         r_dbl_ImpGas = IIf(IsNull(g_rst_Princi!GASADM_PAGFEC), 0, g_rst_Princi!GASADM_PAGFEC) '  ff_GasAdm(g_rst_Princi!SOLMAE_NUMERO, r_dbl_FecPag)
         
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_SOLTRA("
         g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT, "
         g_str_Parame = g_str_Parame & "SOLTRA_FECCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_HORCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_TERCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_NUMSOL, "
         g_str_Parame = g_str_Parame & "SOLTRA_CODOCU, "
         g_str_Parame = g_str_Parame & "SOLTRA_FECING, "
         g_str_Parame = g_str_Parame & "SOLTRA_INGIN1, "
         g_str_Parame = g_str_Parame & "SOLTRA_PAGFEC) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_22.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_DesOcu & "', "
         g_str_Parame = g_str_Parame & CStr(r_dbl_IngAce) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_IngTra) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_FecPag) & ") "
                  
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
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
   If crp_Imprim.Status > 0 Then
      crp_Imprim.RetrieveDataFiles
   End If
   
'   If crp_Imprim.Status > 0 Then
'      crp_Imprim.RetrieveDataFiles
'   End If
   crp_Imprim.Reset
   crp_Imprim.WindowState = crptMaximized
   crp_Imprim.WindowTitle = "Presentación Preliminar"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.WindowShowRefreshBtn = True
   
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se Muestra las tablas que fueron utilizadas en Crystal Report
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_SOLTRA"
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_22.RPT"
   
   crp_Imprim.SelectionFormula = "{RPT_SOLTRA.SOLTRA_NOMRPT} = 'ATE_RPTSOL_22.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SOLTRA.SOLTRA_TERCRE} = '" & modgen_g_str_NombPC & "' "
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Function ff_GasAdm(ByVal p_NumSol As String, Optional ByRef p_FecPag As Double) As Double
   ff_GasAdm = 0
   p_FecPag = 0
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "GASADM_SITUAC = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
         ff_GasAdm = ff_GasAdm + g_rst_Listas!GASADM_PAGIMP
         p_FecPag = g_rst_Listas!GASADM_PAGFEC
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_GenImp_TasSeg()
   Dim r_str_DesOcu     As String
   Dim r_dbl_IngIns     As Double
   Dim r_str_SitTas     As String
   Dim r_str_SitSeg     As String
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
   
   'Se elimina los datos de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT = 'ATE_RPTSOL_23.RPT' AND "
   g_str_Parame = g_str_Parame & "SOLTRA_TERCRE = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
      
   'Leyendo Tabla de solicitudes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_NUMERO, B.PARDES_DESCRI AS ULT_SITUACION, D.SEGUIM_FECINI AS FECING_INSTANCIA, "
   g_str_Parame = g_str_Parame & "        F.SIT_TAS, H.SIT_SEG "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES B ON B.PARDES_CODITE = A.SOLMAE_SITINS AND B.PARDES_CODGRP = '004' "
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_SEGUIM D ON D.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND D.SEGUIM_CODINS = '41'"
   g_str_Parame = g_str_Parame & "         LEFT JOIN ( SELECT F.SEGUIM_NUMSOL, G.PARDES_DESCRI AS SIT_TAS"
   g_str_Parame = g_str_Parame & "                       FROM TRA_SEGUIM F INNER JOIN MNT_PARDES G ON G.PARDES_CODITE = F.SEGUIM_SITUAC AND G.PARDES_CODGRP = '023'"
   g_str_Parame = g_str_Parame & "                      WHERE F.SEGUIM_CODINS = '41'"
   g_str_Parame = g_str_Parame & "                    ) F ON F.SEGUIM_NUMSOL = A.SOLMAE_NUMERO"
   g_str_Parame = g_str_Parame & "         LEFT JOIN ( SELECT H.SEGUIM_NUMSOL, I.PARDES_DESCRI AS SIT_SEG"
   g_str_Parame = g_str_Parame & "                       FROM TRA_SEGUIM H INNER JOIN MNT_PARDES I ON I.PARDES_CODITE = H.SEGUIM_SITUAC AND I.PARDES_CODGRP = '023'"
   g_str_Parame = g_str_Parame & "                      WHERE H.SEGUIM_CODINS = '42') H ON H.SEGUIM_NUMSOL = A.SOLMAE_NUMERO"
   
   g_str_Parame = g_str_Parame & "  WHERE SOLMAE_CODINS = 41 AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
      
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_DesOcu = Trim(g_rst_Princi!ULT_SITUACION) 'moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
         
         'Para obtener Fecha de Ingreso a Instancia
         r_dbl_IngIns = IIf(IsNull(g_rst_Princi!FECING_INSTANCIA), 0, g_rst_Princi!FECING_INSTANCIA) 'ff_IngIns(g_rst_Princi!SOLMAE_NUMERO, 41)
         
         r_str_SitTas = Trim(g_rst_Princi!SIT_TAS) 'ff_SitIns(g_rst_Princi!SOLMAE_NUMERO, 41)
         r_str_SitSeg = Trim(g_rst_Princi!SIT_SEG) 'ff_SitIns(g_rst_Princi!SOLMAE_NUMERO, 42)
         
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_SOLTRA("
         g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT, "
         g_str_Parame = g_str_Parame & "SOLTRA_FECCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_HORCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_TERCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_NUMSOL, "
         g_str_Parame = g_str_Parame & "SOLTRA_CODOCU, "
         g_str_Parame = g_str_Parame & "SOLTRA_FECING, "
         g_str_Parame = g_str_Parame & "SOLTRA_SITIN1, "
         g_str_Parame = g_str_Parame & "SOLTRA_SITIN2) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_23.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_DesOcu & "', "
         g_str_Parame = g_str_Parame & CStr(r_dbl_IngIns) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_SitTas & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_SitSeg & "') "
                  
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
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
   
   crp_Imprim.Reset
   crp_Imprim.WindowState = crptMaximized
   crp_Imprim.WindowTitle = "Presentación Preliminar"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.WindowShowRefreshBtn = True
      
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se Muestra las tablas que fueron utilizadas en Crystal Report
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_SOLTRA"
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_23.RPT"
   
   crp_Imprim.SelectionFormula = "{RPT_SOLTRA.SOLTRA_NOMRPT} = 'ATE_RPTSOL_23.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SOLTRA.SOLTRA_TERCRE} = '" & modgen_g_str_NombPC & "' "
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenImp_EvaLeg()
Dim r_str_DesOcu     As String
Dim r_dbl_FecIng     As Double
        
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
   
   'Se elimina los datos de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT = 'ATE_RPTSOL_24.RPT' AND "
   g_str_Parame = g_str_Parame & "SOLTRA_TERCRE = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
      
   'Leyendo Tabla de solicitudes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_NUMERO, B.PARDES_DESCRI AS SITUACION, C.SEGUIM_FECINI "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES B ON B.PARDES_CODITE = A.SOLMAE_SITINS AND B.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_SEGUIM C ON C.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND C.SEGUIM_CODINS = '21' "
   g_str_Parame = g_str_Parame & "  WHERE SOLMAE_CODINS = 51 AND "
      
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
      
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_DesOcu = Trim(g_rst_Princi!SITUACION) 'moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
                  
         'Para obtener la Fecha de Ingreso a Aprobación Crediticia
         r_dbl_FecIng = IIf(IsNull(g_rst_Princi!SEGUIM_FECINI), 0, g_rst_Princi!SEGUIM_FECINI) 'ff_FecIng(g_rst_Princi!SOLMAE_NUMERO)
         
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
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_24.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_DesOcu & "', "
         g_str_Parame = g_str_Parame & CStr(r_dbl_FecIng) & ") "
                  
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
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
         
   crp_Imprim.Reset
   crp_Imprim.WindowState = crptMaximized
   crp_Imprim.WindowTitle = "Presentación Preliminar"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.WindowShowRefreshBtn = True
   
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se Muestra las tablas que fueron utilizadas en Crystal Report
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_SOLTRA"
   crp_Imprim.DataFiles(4) = "TRA_SEGUIM"
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_24.RPT"
   
   crp_Imprim.SelectionFormula = "{RPT_SOLTRA.SOLTRA_NOMRPT} = 'ATE_RPTSOL_24.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SOLTRA.SOLTRA_TERCRE} = '" & modgen_g_str_NombPC & "' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{TRA_SEGUIM.SEGUIM_CODINS} = 51 "
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenImp_PMVCof()
Dim r_str_DesOcu     As String
Dim r_dbl_IngIns     As Double
Dim r_str_SitTas     As String
Dim r_str_SitSeg     As String
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
   
   'Se elimina los datos de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT = 'ATE_RPTSOL_23.RPT' AND "
   g_str_Parame = g_str_Parame & "SOLTRA_TERCRE = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
      
   'Leyendo Tabla de solicitudes
   'g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_NUMERO, B.PARDES_DESCRI AS ULT_SITUACION, C.SEGUIM_FECINI, D.PARDES_DESCRI AS SIT_TAS, F.PARDES_DESCRI AS SIT_SEG "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES B ON B.PARDES_CODITE = A.SOLMAE_SITINS AND B.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_SEGUIM C ON C.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND C.SEGUIM_CODINS = '41' "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES D ON D.PARDES_CODITE = C.SEGUIM_SITUAC AND D.PARDES_CODGRP = '023' "
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_SEGUIM E ON E.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND E.SEGUIM_CODINS = '42' "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES F ON F.PARDES_CODITE = E.SEGUIM_SITUAC AND F.PARDES_CODGRP = '023' "

   g_str_Parame = g_str_Parame & "  WHERE SOLMAE_CODINS = 41 AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
   
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
      
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_DesOcu = Trim(g_rst_Princi!ULT_SITUACION) 'moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
         
         'Para obtener Fecha de Ingreso a Instancia
         r_dbl_IngIns = IIf(IsNull(g_rst_Princi!SEGUIM_FECINI), 0, g_rst_Princi!SEGUIM_FECINI) 'ff_IngIns(g_rst_Princi!SOLMAE_NUMERO, 41)
         
         r_str_SitTas = Trim(g_rst_Princi!SIT_TAS) 'ff_SitIns(g_rst_Princi!SOLMAE_NUMERO, 41)
         r_str_SitSeg = Trim(g_rst_Princi!SIT_SEG) 'ff_SitIns(g_rst_Princi!SOLMAE_NUMERO, 42)
         
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_SOLTRA("
         g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT, "
         g_str_Parame = g_str_Parame & "SOLTRA_FECCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_HORCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_TERCRE, "
         g_str_Parame = g_str_Parame & "SOLTRA_NUMSOL, "
         g_str_Parame = g_str_Parame & "SOLTRA_CODOCU, "
         g_str_Parame = g_str_Parame & "SOLTRA_FECING, "
         g_str_Parame = g_str_Parame & "SOLTRA_SITIN1, "
         g_str_Parame = g_str_Parame & "SOLTRA_SITIN2) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_23.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_DesOcu & "', "
         g_str_Parame = g_str_Parame & CStr(r_dbl_IngIns) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_SitTas & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_SitSeg & "') "
                  
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
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
   
   crp_Imprim.Reset
   crp_Imprim.WindowState = crptMaximized
   crp_Imprim.WindowTitle = "Presentación Preliminar"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.WindowShowRefreshBtn = True
      
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se Muestra las tablas que fueron utilizadas en Crystal Report
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_SOLTRA"
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_23.RPT"
   
   crp_Imprim.SelectionFormula = "{RPT_SOLTRA.SOLTRA_NOMRPT} = 'ATE_RPTSOL_23.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SOLTRA.SOLTRA_TERCRE} = '" & modgen_g_str_NombPC & "' "
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenExc_AteCli()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PRODUC_DESCRI, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_TIPMON, "
   g_str_Parame = g_str_Parame & "        (TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ' ' || TRIM(DATGEN_NOMBRE)) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "        D.PARDES_DESCRI AS TIPO_EVALUACION, SOLMAE_CONHIP , SOLMAE_FECSOL, SOLMAE_COMVTA_SOL, "
   g_str_Parame = g_str_Parame & "        SOLMAE_COMVTA_DOL, SOLMAE_APOPRO_SOL, SOLMAE_APOPRO_DOL, SOLMAE_MTOPRE_MPR "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES D ON D.PARDES_CODITE = A.SOLMAE_TIPEVA AND D.PARDES_CODGRP = '038'"
   g_str_Parame = g_str_Parame & "  WHERE "
  
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & " SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
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
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!CLIENTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!TIPO_EVALUACION)
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

Private Sub fs_GenExc_EvaCre()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PRODUC_DESCRI, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_TIPMON, SEGUIM_FECINI, "
   g_str_Parame = g_str_Parame & "        (TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ' ' || TRIM(DATGEN_NOMBRE)) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "        E.PARDES_DESCRI AS SITUACION, F.PARDES_DESCRI AS INSTANCIA, SOLMAE_CONHIP , SOLMAE_FECSOL, SOLMAE_COMVTA_SOL, "
   g_str_Parame = g_str_Parame & "        SOLMAE_COMVTA_DOL, SOLMAE_APOPRO_SOL, SOLMAE_APOPRO_DOL, SOLMAE_MTOPRE_MPR "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "  INNER JOIN TRA_SEGUIM D ON D.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND D.SEGUIM_CODINS = 21 "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES E ON E.PARDES_CODITE = D.SEGUIM_SITUAC AND E.PARDES_CODGRP = '023'"
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES F ON F.PARDES_CODITE = A.SOLMAE_SITINS AND F.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "  WHERE A.SOLMAE_CODINS = 21 AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
      
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1  AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 21 "
   g_str_Parame = g_str_Parame & "ORDER BY DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
   
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
      .Cells(1, 7) = "F.INGR. EV. CRED."
      .Cells(1, 8) = "SITUACION EN INSTANCIA"
      .Cells(1, 9) = "ULTIMA OCURRENCIA"
      .Cells(1, 10) = "CONSEJ. HIPOT."
      .Cells(1, 11) = "V. INMUEBLE S/."
      .Cells(1, 12) = "V. INMUEBLE US$"
      .Cells(1, 13) = "PORC. INICIAL"
      .Cells(1, 14) = "MTO. CREDITO S/."
      .Cells(1, 15) = "MTO. CREDITO US$"
      
      .Range(.Cells(1, 1), .Cells(1, 15)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 15)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 30
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 40
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 34
      .Columns("I").ColumnWidth = 26
      .Columns("J").ColumnWidth = 16
      .Columns("K").ColumnWidth = 14
      .Columns("L").ColumnWidth = 14
      .Columns("M").ColumnWidth = 13
      .Columns("N").ColumnWidth = 14
      .Columns("O").ColumnWidth = 14
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!CLIENTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!SITUACION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!INSTANCIA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!SOLMAE_CONHIP)
            
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
      End If
      
      If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Or g_rst_Princi!SOLMAE_COMVTA_DOL > 0 Then
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_SOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00") * 100) + "%"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_DOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00") * 100) + "%"
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(0) + "%"
      End If
            
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
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

Private Sub fs_GenExc_AceCre()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_dbl_IngAce     As Double
   Dim r_dbl_IngTra     As Double
   Dim r_dbl_ImpGas     As Double
   Dim r_dbl_FecPag     As Double

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PRODUC_DESCRI, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_TIPMON, "
   g_str_Parame = g_str_Parame & "        (TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ' ' || TRIM(DATGEN_NOMBRE)) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "        D.PARDES_DESCRI AS SITUACION, SOLMAE_CONHIP , SOLMAE_FECSOL, SOLMAE_COMVTA_SOL, "
   g_str_Parame = g_str_Parame & "        SOLMAE_COMVTA_DOL, SOLMAE_APOPRO_SOL, SOLMAE_APOPRO_DOL, SOLMAE_MTOPRE_MPR, NVL(E.PAGIMP,0) AS PAGIMP,"
   g_str_Parame = g_str_Parame & "        E.GASADM_PAGFEC , F.SEGUIM_FECINI AS FECINI_ACECRED, G.SEGUIM_FECINI AS FECINI_TRACLI "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES D ON D.PARDES_CODITE = A.SOLMAE_SITINS AND D.PARDES_CODGRP = '004' "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT GASADM_NUMSOL, SUM(GASADM_PAGIMP) AS PAGIMP, GASADM_PAGFEC FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "                           GASADM_SITUAC = 1  GROUP BY GASADM_NUMSOL, GASADM_PAGFEC) E ON E.GASADM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT SEGUIM_NUMSOL, SEGUIM_FECINI FROM TRA_SEGUIM F WHERE F.SEGUIM_CODINS = '31') F ON F.SEGUIM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT SEGUIM_NUMSOL, SEGUIM_FECINI FROM TRA_SEGUIM G WHERE G.SEGUIM_CODINS = '31') G ON G.SEGUIM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_SEGUIM H ON H.SEGUIM_NUMSOL = A.SOLMAE_NUMERO"
   g_str_Parame = g_str_Parame & "  WHERE (A.SOLMAE_CODINS = 31 OR A.SOLMAE_CODINS = 32) AND "
   
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "A.SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
      
   g_str_Parame = g_str_Parame & "A.SOLMAE_SITUAC = 1  AND "
   g_str_Parame = g_str_Parame & "H.SEGUIM_CODINS = 21 "
   g_str_Parame = g_str_Parame & "ORDER BY B.DATGEN_APEPAT ASC, B.DATGEN_APEMAT ASC, B.DATGEN_NOMBRE ASC"
   
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
      .Cells(1, 7) = "F. INGR. ACEPT. CRED."
      .Cells(1, 8) = "F. INGR. TRAMITES"
      .Cells(1, 9) = "F. PAGO GC"
      
      .Cells(1, 10) = "ULTIMA OCURRENCIA"
      .Cells(1, 11) = "CONSEJ. HIPOT."
      .Cells(1, 12) = "V. INMUEBLE S/."
      .Cells(1, 13) = "V. INMUEBLE US$"
      .Cells(1, 14) = "PORC. INICIAL"
      .Cells(1, 15) = "MTO. CREDITO S/."
      .Cells(1, 16) = "MTO. CREDITO US$"
      
      .Range(.Cells(1, 1), .Cells(1, 16)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 16)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      
      .Columns("B").ColumnWidth = 30
      
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 40
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 15
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 15
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 34
      .Columns("K").ColumnWidth = 26
      .Columns("L").ColumnWidth = 16
      
      .Columns("M").ColumnWidth = 14
      .Columns("N").ColumnWidth = 14
      .Columns("O").ColumnWidth = 13
      .Columns("P").ColumnWidth = 14
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI) 'moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!CLIENTE) 'moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      
      'Para obtener Fecha de Ingreso a Aceptación Crediticia
      r_dbl_IngAce = ff_IngIns(g_rst_Princi!SOLMAE_NUMERO, 31) 'IIf(IsNull(g_rst_Princi!FECINI_ACECRED), 0, g_rst_Princi!FECINI_ACECRED) '
      
      'Para obtener Fecha de Ingreso a Trámites de Cliente
      r_dbl_IngTra = ff_IngIns(g_rst_Princi!SOLMAE_NUMERO, 32) 'IIf(IsNull(g_rst_Princi!FECINI_TRACLI), 0, g_rst_Princi!FECINI_TRACLI)  '
      
      'Para obtener Fecha de Pago de Gastos de Cierre
      'r_dbl_ImpGas = ff_GasAdm(g_rst_Princi!SOLMAE_NUMERO, r_dbl_FecPag)
      r_dbl_ImpGas = g_rst_Princi!PAGIMP
      r_dbl_FecPag = IIf(IsNull(g_rst_Princi!GASADM_PAGFEC), 0, g_rst_Princi!GASADM_PAGFEC)
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(r_dbl_IngAce)))
      
      If r_dbl_IngTra > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(CStr(r_dbl_IngTra)))
      End If
      
      If r_dbl_FecPag > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CDate(gf_FormatoFecha(CStr(r_dbl_FecPag)))
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!SITUACION) 'moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!SOLMAE_CONHIP)
            
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
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
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
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

Private Sub fs_GenExc_TasSeg()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_dbl_IngIns     As Double
   Dim r_str_SitTas     As String
   Dim r_str_SitSeg     As String

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PRODUC_DESCRI, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_TIPMON, "
   g_str_Parame = g_str_Parame & "        (TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ' ' || TRIM(DATGEN_NOMBRE)) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "        D.PARDES_DESCRI AS INSTANCIA, SOLMAE_CONHIP , SOLMAE_FECSOL, SOLMAE_COMVTA_SOL, "
   g_str_Parame = g_str_Parame & "        SOLMAE_COMVTA_DOL, SOLMAE_APOPRO_SOL, SOLMAE_APOPRO_DOL, SOLMAE_MTOPRE_MPR, "
   g_str_Parame = g_str_Parame & "        F.SIT_TAS , H.SIT_SEG, E.SEGUIM_FECINI "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON D.PARDES_CODITE = A.SOLMAE_SITINS AND D.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "  INNER JOIN TRA_SEGUIM E ON E.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND E.SEGUIM_CODINS = 41"
   g_str_Parame = g_str_Parame & "   LEFT JOIN ( SELECT F.SEGUIM_NUMSOL, G.PARDES_DESCRI AS SIT_TAS"
   g_str_Parame = g_str_Parame & "                 FROM TRA_SEGUIM F INNER JOIN MNT_PARDES G ON G.PARDES_CODITE = F.SEGUIM_SITUAC AND G.PARDES_CODGRP = '023'"
   g_str_Parame = g_str_Parame & "                WHERE F.SEGUIM_CODINS = '41'"
   g_str_Parame = g_str_Parame & "                    ) F ON F.SEGUIM_NUMSOL = A.SOLMAE_NUMERO"
   g_str_Parame = g_str_Parame & "   LEFT JOIN ( SELECT H.SEGUIM_NUMSOL, I.PARDES_DESCRI AS SIT_SEG"
   g_str_Parame = g_str_Parame & "                 FROM TRA_SEGUIM H INNER JOIN MNT_PARDES I ON I.PARDES_CODITE = H.SEGUIM_SITUAC AND I.PARDES_CODGRP = '023'"
   g_str_Parame = g_str_Parame & "                WHERE H.SEGUIM_CODINS = '42') H ON H.SEGUIM_NUMSOL = A.SOLMAE_NUMERO"
   g_str_Parame = g_str_Parame & "  WHERE A.SOLMAE_CODINS = 41 AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
      
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1  AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 41 "
   g_str_Parame = g_str_Parame & "ORDER BY DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
   
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
      .Cells(1, 7) = "F. INGR. EVALUACION"
      .Cells(1, 8) = "SITUACION TASACION"
      .Cells(1, 9) = "SITUACION SEGUROS"
      
      .Cells(1, 10) = "ULTIMA OCURRENCIA"
      .Cells(1, 11) = "CONSEJ. HIPOT."
      .Cells(1, 12) = "V. INMUEBLE S/."
      .Cells(1, 13) = "V. INMUEBLE US$"
      .Cells(1, 14) = "PORC. INICIAL"
      .Cells(1, 15) = "MTO. CREDITO S/."
      .Cells(1, 16) = "MTO. CREDITO US$"
      
      .Range(.Cells(1, 1), .Cells(1, 16)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 16)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 30
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 40
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 15
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 15
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 34
      .Columns("K").ColumnWidth = 26
      .Columns("L").ColumnWidth = 16
      .Columns("M").ColumnWidth = 14
      .Columns("N").ColumnWidth = 14
      .Columns("O").ColumnWidth = 13
      .Columns("P").ColumnWidth = 14
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!CLIENTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      
      'Para obtener Fecha de Ingreso a Instancia
      r_dbl_IngIns = IIf(IsNull(g_rst_Princi!SEGUIM_FECINI), 0, g_rst_Princi!SEGUIM_FECINI) 'ff_IngIns(g_rst_Princi!SOLMAE_NUMERO, 41)
      
      r_str_SitTas = Trim(g_rst_Princi!SIT_TAS)
      r_str_SitSeg = Trim(g_rst_Princi!SIT_SEG)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(r_dbl_IngIns)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = r_str_SitTas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = r_str_SitSeg
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!INSTANCIA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!SOLMAE_CONHIP)
            
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
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
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
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

Private Sub fs_GenExc_EvaLeg()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT TRIM(PRODUC_DESCRI) AS PRODUC_DESCRI, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_TIPMON, SOLMAE_FECSOL, "
   g_str_Parame = g_str_Parame & "        (TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ' ' || TRIM(DATGEN_NOMBRE)) AS CLIENTE, C.SEGUIM_FECINI, "
   g_str_Parame = g_str_Parame & "        TRIM(E.PARDES_DESCRI) AS SITUACION, TRIM(F.PARDES_DESCRI) AS INSTANCIA, TRIM(SOLMAE_CONHIP) AS SOLMAE_CONHIP, SOLMAE_COMVTA_SOL, "
   g_str_Parame = g_str_Parame & "        SOLMAE_COMVTA_DOL, SOLMAE_APOPRO_SOL, SOLMAE_APOPRO_DOL, SOLMAE_MTOPRE_MPR "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A"
   g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.SOLMAE_TITTDO AND B.DATGEN_NUMDOC = A.SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & "  INNER JOIN TRA_SEGUIM C ON C.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND C.SEGUIM_CODINS = 51"
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_PRODUC D ON D.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES E ON E.PARDES_CODITE = C.SEGUIM_SITUAC AND E.PARDES_CODGRP = '023'"
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES F ON F.PARDES_CODITE = A.SOLMAE_SITINS AND F.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "  WHERE SOLMAE_CODINS = 51 AND "
  
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
      
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1  AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 51 "
   g_str_Parame = g_str_Parame & "ORDER BY DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
   
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
      .Cells(1, 7) = "F.INGR. EVALUAC."
      .Cells(1, 8) = "SITUACION EN INSTANCIA"
      .Cells(1, 9) = "ULTIMA OCURRENCIA"
      .Cells(1, 10) = "CONSEJ. HIPOT."
      .Cells(1, 11) = "V. INMUEBLE S/."
      .Cells(1, 12) = "V. INMUEBLE US$"
      .Cells(1, 13) = "PORC. INICIAL"
      .Cells(1, 14) = "MTO. CREDITO S/."
      .Cells(1, 15) = "MTO. CREDITO US$"
      
      .Range(.Cells(1, 1), .Cells(1, 15)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 15)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 30
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 40
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 34
      .Columns("I").ColumnWidth = 26
      .Columns("J").ColumnWidth = 16
      .Columns("K").ColumnWidth = 14
      .Columns("L").ColumnWidth = 14
      .Columns("M").ColumnWidth = 13
      .Columns("N").ColumnWidth = 14
      .Columns("O").ColumnWidth = 14
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!CLIENTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!SITUACION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!INSTANCIA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
      End If
      If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Or g_rst_Princi!SOLMAE_COMVTA_DOL > 0 Then
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_SOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00") * 100) + "%"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_DOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00") * 100) + "%"
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(0) + "%"
      End If
            
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
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

Private Sub fs_GenExc_PMVCof()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_dbl_IngIns     As Double
   Dim r_str_SitTas     As String
   Dim r_str_SitSeg     As String
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PRODUC_DESCRI, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_TIPMON, "
   g_str_Parame = g_str_Parame & "        (TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ' ' || TRIM(DATGEN_NOMBRE)) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "        SOLMAE_CONHIP , SOLMAE_FECSOL, SOLMAE_COMVTA_SOL, SOLMAE_COMVTA_DOL, "
   g_str_Parame = g_str_Parame & "        SOLMAE_APOPRO_SOL, SOLMAE_APOPRO_DOL, SOLMAE_MTOPRE_MPR, "
   g_str_Parame = g_str_Parame & "        D.SEGUIM_FECINI, E.PARDES_DESCRI AS SIT_TAS, SIT_SEG,  G.PARDES_DESCRI AS OCURRENCIA "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A"
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_SEGUIM D ON D.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND D.SEGUIM_CODINS = '41'"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES E ON E.PARDES_CODITE = D.SEGUIM_SITUAC AND E.PARDES_CODGRP = '023'"
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT F.SEGUIM_NUMSOL, G.PARDES_DESCRI AS SIT_SEG "
   g_str_Parame = g_str_Parame & "                      FROM TRA_SEGUIM F "
   g_str_Parame = g_str_Parame & "                           INNER JOIN MNT_PARDES G ON G.PARDES_CODITE = F.SEGUIM_SITUAC AND G.PARDES_CODGRP = '023' "
   g_str_Parame = g_str_Parame & "                     WHERE F.SEGUIM_CODINS = '42') F ON F.SEGUIM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES G ON G.PARDES_CODITE = A.SOLMAE_SITINS AND G.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "  WHERE SOLMAE_CODINS = 41 AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
      
   g_str_Parame = g_str_Parame & "A.SOLMAE_SITUAC = 1  AND "
   g_str_Parame = g_str_Parame & "D.SEGUIM_CODINS = 41 "
   g_str_Parame = g_str_Parame & "ORDER BY B.DATGEN_APEPAT ASC, B.DATGEN_APEMAT ASC, B.DATGEN_NOMBRE ASC"
   
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
      .Cells(1, 7) = "F. INGR. EVALUACION"
      .Cells(1, 8) = "SITUACION TASACION"
      .Cells(1, 9) = "SITUACION SEGUROS"
      
      .Cells(1, 10) = "ULTIMA OCURRENCIA"
      .Cells(1, 11) = "CONSEJ. HIPOT."
      .Cells(1, 12) = "V. INMUEBLE S/."
      .Cells(1, 13) = "V. INMUEBLE US$"
      .Cells(1, 14) = "PORC. INICIAL"
      .Cells(1, 15) = "MTO. CREDITO S/."
      .Cells(1, 16) = "MTO. CREDITO US$"
      
      .Range(.Cells(1, 1), .Cells(1, 16)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 16)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      
      .Columns("B").ColumnWidth = 30
      
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 40
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 15
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 15
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 34
      .Columns("K").ColumnWidth = 26
      .Columns("L").ColumnWidth = 16
      
      .Columns("M").ColumnWidth = 14
      .Columns("N").ColumnWidth = 14
      .Columns("O").ColumnWidth = 13
      .Columns("P").ColumnWidth = 14
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI) 'moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!CLIENTE) 'moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      
      'Para obtener Fecha de Ingreso a Instancia
      r_dbl_IngIns = IIf(IsNull(g_rst_Princi!SEGUIM_FECINI), 0, g_rst_Princi!SEGUIM_FECINI) 'ff_IngIns(g_rst_Princi!SOLMAE_NUMERO, 41)
      
      r_str_SitTas = Trim(g_rst_Princi!SIT_TAS) 'ff_SitIns(g_rst_Princi!SOLMAE_NUMERO, 41)
      r_str_SitSeg = Trim(g_rst_Princi!SIT_SEG) 'ff_SitIns(g_rst_Princi!SOLMAE_NUMERO, 42)
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(r_dbl_IngIns)))
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = r_str_SitTas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = r_str_SitSeg
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!OCURRENCIA) 'moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
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
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
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
