VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_45 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7680
   Icon            =   "AteCli_frm_571.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   4275
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7695
      _Version        =   65536
      _ExtentX        =   13573
      _ExtentY        =   7541
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   7605
         _Version        =   65536
         _ExtentX        =   13414
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
            Left            =   6990
            Picture         =   "AteCli_frm_571.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_571.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_571.frx":0890
            Style           =   1  'Graphical
            TabIndex        =   7
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   7605
         _Version        =   65536
         _ExtentX        =   13414
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
            TabIndex        =   12
            Top             =   30
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes en Tramite x Proyecto"
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
         Begin Threed.SSPanel ssp_TipCon 
            Height          =   285
            Left            =   660
            TabIndex        =   13
            Top             =   330
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   503
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_571.frx":0B9A
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2055
         Left            =   30
         TabIndex        =   14
         Top             =   2160
         Width           =   7605
         _Version        =   65536
         _ExtentX        =   13414
         _ExtentY        =   3625
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
         Begin VB.CheckBox chk_TipPry 
            Caption         =   "Todos los Tipos de Proyectos"
            Height          =   285
            Left            =   1530
            TabIndex        =   2
            Top             =   390
            Width           =   2475
         End
         Begin VB.CheckBox chk_TipCon 
            Caption         =   "Todos"
            Height          =   315
            Left            =   1530
            TabIndex        =   5
            Top             =   1710
            Width           =   4515
         End
         Begin VB.CheckBox chk_Proyec 
            Caption         =   "Todos los Proyectos"
            Height          =   285
            Left            =   1530
            TabIndex        =   3
            Top             =   1050
            Width           =   1845
         End
         Begin VB.ComboBox cmb_TipPry 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   6015
         End
         Begin VB.ComboBox cmb_Proyec 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   720
            Width           =   6015
         End
         Begin VB.ComboBox cmb_TipCon 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1380
            Width           =   6015
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Proyecto:"
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Proyecto:"
            Height          =   255
            Left            =   60
            TabIndex        =   17
            Top             =   720
            Width           =   885
         End
         Begin VB.Label lbl_TipCon 
            Caption         =   "Tipo de Consulta:"
            Height          =   435
            Left            =   60
            TabIndex        =   16
            Top             =   1380
            Width           =   1275
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   675
         Left            =   30
         TabIndex        =   19
         Top             =   1440
         Width           =   8445
         _Version        =   65536
         _ExtentX        =   14896
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
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   6015
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Reporte:"
            Height          =   255
            Left            =   60
            TabIndex        =   20
            Top             =   210
            Width           =   1290
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_45"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()   As moddat_tpo_Genera
Dim l_arr_ConHip()   As moddat_tpo_Genera
Dim l_arr_Proyec()   As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub chk_TipCon_Click()
   If chk_TipCon.Value = 1 Then
      cmb_TipCon.ListIndex = -1
      cmb_TipCon.Enabled = False
      Call gs_SetFocus(cmd_Imprim)
   ElseIf chk_TipCon.Value = 0 Then
      cmb_TipCon.Enabled = True
      Call gs_SetFocus(cmb_TipCon)
   End If
End Sub

Private Sub chk_Proyec_Click()
   If chk_Proyec.Value = 1 Then
      cmb_Proyec.ListIndex = -1
      cmb_Proyec.Enabled = False
      Call gs_SetFocus(cmb_TipCon)
   ElseIf chk_Proyec.Value = 0 Then
      cmb_Proyec.Enabled = True
      Call gs_SetFocus(cmb_Proyec)
   End If
End Sub

Private Sub chk_TipPry_Click()
   If chk_TipPry.Value = 1 Then
      cmb_TipPry.ListIndex = -1
      cmb_TipPry.Enabled = False
      cmb_Proyec.Enabled = False
      chk_Proyec.Value = 1
      chk_Proyec.Enabled = False
      Call gs_SetFocus(cmb_TipCon)
   ElseIf chk_TipPry.Value = 0 Then
      chk_Proyec.Enabled = True
      cmb_TipPry.Enabled = True
      cmb_Proyec.Enabled = True
      chk_Proyec.Value = 0
      
      Call gs_SetFocus(cmb_TipPry)
   End If
End Sub
  
Private Sub cmb_TipPry_Click()
   
   If cmb_TipPry.ListIndex > -1 Then
      Screen.MousePointer = 11
      
      Call Carga_PryInm_Combo(cmb_Proyec, l_arr_Proyec, cmb_TipPry.ItemData(cmb_TipPry.ListIndex))
      
      Screen.MousePointer = 0
   End If
   
End Sub

Private Sub cmb_TipPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPry_Click
   End If
End Sub


Private Sub cmb_TipRep_Click()
   Call fs_Limpia
   If cmb_TipRep.ListIndex <> -1 Then
      If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
         Call moddat_gs_Carga_Produc(cmb_TipCon, l_arr_Produc, 4)
         chk_TipCon.Caption = "Todos los Productos"
         lbl_TipCon.Caption = "Producto:"
         ssp_TipCon.Caption = "Por Producto"
      Else
         Call moddat_gs_Carga_EjecMC(cmb_TipCon, l_arr_ConHip, 121)
         chk_TipCon.Caption = "Todos los Consejeros Hipotecarios"
         lbl_TipCon.Caption = "Consejero Hipotecario:"
         ssp_TipCon.Caption = "Por Consejero Hipotecario"
      End If
   End If
   Call gs_SetFocus(cmb_TipPry)
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipRep_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
    
   'Validación
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   If chk_Proyec = 0 Then
      If cmb_Proyec.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Proyec)
         Exit Sub
      End If
   End If
   
   If chk_TipPry = 0 Then
      If cmb_TipPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Tipo de Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPry)
         Exit Sub
      End If
   End If
   
   If chk_TipCon = 0 Then
      If cmb_TipCon.ListIndex = -1 Then
         If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
            MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         End If
         Call gs_SetFocus(cmb_TipCon)
         Exit Sub
      End If
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      Call fs_GenExc_TipPro
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      Call fs_GenExc_ConHip
   End If
End Sub

Private Sub cmd_Imprim_Click()
   
   'Validación
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   If chk_Proyec = 0 Then
      If cmb_Proyec.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Proyec)
         Exit Sub
      End If
   End If
   
   If chk_TipPry = 0 Then
      If cmb_TipPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Tipo de Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPry)
         Exit Sub
      End If
   End If
   
   If chk_TipCon = 0 Then
      If cmb_TipCon.ListIndex = -1 Then
         If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
            MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         End If
         Call gs_SetFocus(cmb_TipCon)
         Exit Sub
      End If
   End If
   
   'Confirmación
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      Call fs_GenImp_TipPro
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      Call fs_GenImp_ConHip
   End If
   
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   Call fs_Limpia
   
   cmb_TipRep.AddItem "POR PRODUCTO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
      
   If modgen_g_int_TipUsu <> 20121 Then
      cmb_TipRep.AddItem "POR CONSEJERO HIPOTECARIO"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2
   End If
   
   cmb_TipRep.ListIndex = -1
   
   cmb_TipPry.AddItem ("VINCULADO")
   cmb_TipPry.ItemData(cmb_TipPry.NewIndex) = 1
      
   cmb_TipPry.AddItem ("NO VINCULADO")
   cmb_TipPry.ItemData(cmb_TipPry.NewIndex) = 2
End Sub

Private Sub Carga_PryInm_Combo(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipPry As Integer)
   ReDim p_Arregl(0)
   p_Combo.Clear
      
   g_str_Parame = "SELECT DATGEN_CODIGO, DATGEN_TITULO FROM PRY_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_PRYMCS = " & CStr(p_TipPry) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY DATGEN_TITULO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
      
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!DATGEN_TITULO)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!DATGEN_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!DATGEN_TITULO)
            
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub
Private Sub fs_GenImp_ConHip()
   Dim r_str_PryMcs     As String
   Dim r_str_DesOcu     As String
   Dim r_str_DesIns     As String
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_DESPRY WHERE "
   g_str_Parame = g_str_Parame & "DESPRY_NOMRPT = 'ATE_RPTSOL_10.RPT' AND "
   g_str_Parame = g_str_Parame & "DESPRY_TERCRE = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   'g_str_Parame = "SELECT * FROM CRE_SOLMAE, PRY_DATGEN, CRE_SOLINM WHERE "
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_NUMERO, C.DATGEN_TITULO, D.PARDES_DESCRI AS INSTANCIA, E.PARDES_DESCRI AS SITUACION, F.PARDES_DESCRI AS PROYECTO "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_SOLINM B ON A.SOLMAE_NUMERO = B.SOLINM_NUMSOL "
   g_str_Parame = g_str_Parame & "        INNER JOIN PRY_DATGEN C ON C.DATGEN_CODIGO = B.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES D ON D.PARDES_CODITE = A.SOLMAE_CODINS AND D.PARDES_CODGRP = '002'"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES E ON E.PARDES_CODITE = A.SOLMAE_SITINS AND E.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES F ON F.PARDES_CODITE = C.DATGEN_PRYMCS AND F.PARDES_CODGRP = '214'"
   g_str_Parame = g_str_Parame & "  WHERE "
   
   'Si no escogio todos los Productos
   If chk_TipPry.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_PRYMCS = '" & (cmb_TipPry.ListIndex + 1) & "' AND "
   End If
   
   If chk_Proyec.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_CODIGO = '" & l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLINM_NUMSOL AND "
   g_str_Parame = g_str_Parame & "DATGEN_CODIGO = SOLINM_PRYCOD AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
      g_rst_Princi.MoveFirst
   
      Do While Not g_rst_Princi.EOF
            
         'Para obtener SI es un proyecto vinculado (Mi Casita)
         r_str_PryMcs = Trim(g_rst_Princi!PROYECTO) 'moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_PRYMCS))
         
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_DesOcu = Trim(g_rst_Princi!SITUACION) 'moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
         
         'Para obtener Descripción de Instancia Actual
         r_str_DesIns = Trim(g_rst_Princi!INSTANCIA) 'moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SOLMAE_CODINS))
         
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_DESPRY("
         g_str_Parame = g_str_Parame & "DESPRY_NOMRPT, "
         g_str_Parame = g_str_Parame & "DESPRY_FECCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_HORCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_TERCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_NUMSOL, "
         g_str_Parame = g_str_Parame & "DESPRY_PRYTIT, "
         g_str_Parame = g_str_Parame & "DESPRY_PRYMCS, "
         g_str_Parame = g_str_Parame & "DESPRY_CODOCU) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_10.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DATGEN_TITULO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_PryMcs & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_DesOcu & "') "
      
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
  
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_DESPRY"
   
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_10.RPT"
   
   crp_Imprim.SelectionFormula = "{RPT_DESPRY.DESPRY_NOMRPT} = 'ATE_RPTSOL_10.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_DESPRY.DESPRY_TERCRE} = '" & modgen_g_str_NombPC & "' "
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenExc_ConHip()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PRODUC_DESCRI, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_TIPMON, D.DATGEN_TITULO, D.DATGEN_PRYMCS, "
   g_str_Parame = g_str_Parame & "        (TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ' ' || TRIM(DATGEN_NOMBRE)) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "        F.PARDES_DESCRI AS INSTANCIA, G.PARDES_DESCRI AS SITUACION, SOLMAE_CONHIP , SOLMAE_FECSOL, SOLMAE_COMVTA_SOL, "
   g_str_Parame = g_str_Parame & "        SOLMAE_COMVTA_DOL, SOLMAE_APOPRO_SOL, SOLMAE_APOPRO_DOL, SOLMAE_MTOPRE_MPR "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A"
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_SOLINM C ON A.SOLMAE_NUMERO = C.SOLINM_NUMSOL "
   g_str_Parame = g_str_Parame & "        INNER JOIN PRY_DATGEN D ON D.DATGEN_CODIGO = C.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC E ON E.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES F ON F.PARDES_CODITE = A.SOLMAE_CODINS AND F.PARDES_CODGRP = '002'"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES G ON G.PARDES_CODITE = A.SOLMAE_SITINS AND G.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "  WHERE "
   
   'Si no escogio todos los Productos
   If chk_TipPry.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_PRYMCS = '" & (cmb_TipPry.ListIndex + 1) & "' AND "
   End If
   
   If chk_Proyec.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_CODIGO = '" & l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLINM_NUMSOL AND "
   g_str_Parame = g_str_Parame & "DATGEN_CODIGO = SOLINM_PRYCOD AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = DATGEN_TIPDOC AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = DATGEN_NUMDOC AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CONHIP ASC, DATGEN_TITULO ASC, DATGEN_PRYMCS ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
          
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
      
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "CONSEJ. HIPOT."
      .Cells(1, 3) = "PROYECTO"
      .Cells(1, 4) = "TIPO DE PROYECTO"
      .Cells(1, 5) = "PRODUCTO"
      .Cells(1, 6) = "SOLICITUD"
      .Cells(1, 7) = "DOC. IDENTIDAD"
      .Cells(1, 8) = "NOMBRE CLIENTE"
      .Cells(1, 9) = "F. SOLICITUD"
      .Cells(1, 10) = "INSTANCIA ACTUAL"
      .Cells(1, 11) = "SITUAC. INSTANCIA"
      .Cells(1, 12) = "TIP. MONEDA"
      .Cells(1, 13) = "V. INMUEBLE S/."
      .Cells(1, 14) = "V. INMUEBLE US$."
      .Cells(1, 15) = "PORC. INICIAL"
      .Cells(1, 16) = "MTO. CREDITO S/."
      .Cells(1, 17) = "MTO. CREDITO US$."
      
      .Range(.Cells(1, 1), .Cells(1, 17)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 17)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 15
      .Columns("C").ColumnWidth = 50
      
      .Columns("D").ColumnWidth = 50
      
      .Columns("E").ColumnWidth = 34
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 16
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 39
            
      .Columns("I").ColumnWidth = 15
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 31
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 38
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 23
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      
      
      .Columns("M").ColumnWidth = 17
      .Columns("N").ColumnWidth = 13
      .Columns("O").ColumnWidth = 17
      .Columns("P").ColumnWidth = 18
      .Columns("Q").ColumnWidth = 18
   End With
   
   g_rst_Princi.MoveFirst
     
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!DATGEN_TITULO)
      
      If g_rst_Princi!DATGEN_PRYMCS = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "PROYECTO VINCULADO"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "PROYECTO NO VINCULADO"
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!PRODUC_DESCRI) 'moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!CLIENTE) 'moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!INSTANCIA) 'moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SOLMAE_CODINS))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!SITUACION) 'moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
               
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = "SOLES"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = "DOLARES AMERICANOS"
      End If
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = 0
      End If
         
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = 0
      End If
              
      If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Or g_rst_Princi!SOLMAE_COMVTA_DOL > 0 Then
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!SOLMAE_APOPRO_SOL / g_rst_Princi!SOLMAE_COMVTA_SOL * 100, "##0.00") + "%"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!SOLMAE_APOPRO_DOL / g_rst_Princi!SOLMAE_COMVTA_DOL * 100, "##0.00") + "%"
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = "0.00%"
      End If
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
      End If
         
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = 0
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
Private Sub fs_GenImp_TipPro()
   Dim r_str_PryMcs     As String
   Dim r_str_DesOcu     As String
   Dim r_str_DesIns     As String
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_DESPRY WHERE "
   g_str_Parame = g_str_Parame & "DESPRY_NOMRPT = 'ATE_RPTSOL_09.RPT' AND "
   g_str_Parame = g_str_Parame & "DESPRY_TERCRE = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   'g_str_Parame = "SELECT * FROM CRE_SOLMAE, PRY_DATGEN, CRE_SOLINM WHERE "
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_NUMERO, C.DATGEN_TITULO, D.PARDES_DESCRI AS INSTANCIA, E.PARDES_DESCRI AS SITUACION, F.PARDES_DESCRI AS PROYECTO "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_SOLINM B ON A.SOLMAE_NUMERO = B.SOLINM_NUMSOL "
   g_str_Parame = g_str_Parame & "        INNER JOIN PRY_DATGEN C ON C.DATGEN_CODIGO = B.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES D ON D.PARDES_CODITE = A.SOLMAE_CODINS AND D.PARDES_CODGRP = '002'"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES E ON E.PARDES_CODITE = A.SOLMAE_SITINS AND E.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES F ON F.PARDES_CODITE = C.DATGEN_PRYMCS AND F.PARDES_CODGRP = '214'"
   g_str_Parame = g_str_Parame & "  WHERE "
   
   'Si no escogio todos los Productos
   If chk_TipPry.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_PRYMCS = '" & (cmb_TipPry.ListIndex + 1) & "' AND "
   End If
   
   If chk_Proyec.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_CODIGO = '" & l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLINM_NUMSOL AND "
   g_str_Parame = g_str_Parame & "DATGEN_CODIGO = SOLINM_PRYCOD AND "
   
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
            
         'Para obtener SI es un proyecto vinculado (Mi Casita)
         r_str_PryMcs = Trim(g_rst_Princi!PROYECTO) 'moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_PRYMCS))
         
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_DesOcu = Trim(g_rst_Princi!SITUACION) 'moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
         
         'Para obtener Descripción de Instancia Actual
         r_str_DesIns = Trim(g_rst_Princi!INSTANCIA) 'moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SOLMAE_CODINS))
         
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_DESPRY("
         g_str_Parame = g_str_Parame & "DESPRY_NOMRPT, "
         g_str_Parame = g_str_Parame & "DESPRY_FECCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_HORCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_TERCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_NUMSOL, "
         g_str_Parame = g_str_Parame & "DESPRY_PRYTIT, "
         g_str_Parame = g_str_Parame & "DESPRY_PRYMCS, "
         g_str_Parame = g_str_Parame & "DESPRY_CODOCU) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_09.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DATGEN_TITULO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_PryMcs & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_DesOcu & "') "
               
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
  
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_DESPRY"
   
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_09.RPT"
   
   crp_Imprim.SelectionFormula = "{RPT_DESPRY.DESPRY_NOMRPT} = 'ATE_RPTSOL_09.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_DESPRY.DESPRY_TERCRE} = '" & modgen_g_str_NombPC & "' "
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenExc_TipPro()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PRODUC_DESCRI, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_TIPMON, DATGEN_TITULO, DATGEN_PRYMCS, "
   g_str_Parame = g_str_Parame & "        (TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ' ' || TRIM(DATGEN_NOMBRE)) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "        F.PARDES_DESCRI AS INSTANCIA, G.PARDES_DESCRI AS SITUACION, SOLMAE_CONHIP , SOLMAE_FECSOL, SOLMAE_COMVTA_SOL, "
   g_str_Parame = g_str_Parame & "        SOLMAE_COMVTA_DOL, SOLMAE_APOPRO_SOL, SOLMAE_APOPRO_DOL, SOLMAE_MTOPRE_MPR "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A"
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_SOLINM C ON A.SOLMAE_NUMERO = C.SOLINM_NUMSOL "
   g_str_Parame = g_str_Parame & "        INNER JOIN PRY_DATGEN D ON D.DATGEN_CODIGO = C.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC E ON E.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES F ON F.PARDES_CODITE = A.SOLMAE_CODINS AND F.PARDES_CODGRP = '002'"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES G ON G.PARDES_CODITE = A.SOLMAE_SITINS AND G.PARDES_CODGRP = '004'"
   g_str_Parame = g_str_Parame & "  WHERE "
   
   If chk_TipPry.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_PRYMCS = '" & (cmb_TipPry.ListIndex + 1) & "' AND "
   End If
   
   If chk_Proyec.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_CODIGO = '" & l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_Codigo & "' AND "
   End If
      
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
    
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY DATGEN_TITULO ASC, DATGEN_PRYMCS ASC, SOLMAE_CODPRD ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
          
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
      
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "PROYECTO"
      .Cells(1, 3) = "TIPO PROYECTO"
      .Cells(1, 4) = "PRODUCTO"
      .Cells(1, 5) = "SOLICITUD"
      .Cells(1, 6) = "DOC. IDENTIDAD"
      .Cells(1, 7) = "NOMBRE CLIENTE"
      .Cells(1, 8) = "F. SOLICITUD"
      .Cells(1, 9) = "INSTANCIA ACTUAL"
      .Cells(1, 10) = "SITUAC. INSTANCIA"
      .Cells(1, 11) = "CONSEJ. HIPOT."
      .Cells(1, 12) = "TIP. MONEDA"
      .Cells(1, 13) = "V. INMUEBLE S/."
      .Cells(1, 14) = "V. INMUEBLE US$."
      .Cells(1, 15) = "PORC. INICIAL"
      .Cells(1, 16) = "MTO. CREDITO S/."
      .Cells(1, 17) = "MTO. CREDITO US$."
      
      .Range(.Cells(1, 1), .Cells(1, 17)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 17)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 50
      .Columns("C").ColumnWidth = 50
      
      .Columns("D").ColumnWidth = 32
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 40
      
      .Columns("H").ColumnWidth = 12
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 31
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 50
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 22
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 21
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      
      .Columns("M").ColumnWidth = 17
      .Columns("N").ColumnWidth = 13
      .Columns("O").ColumnWidth = 17
      .Columns("P").ColumnWidth = 18
      .Columns("Q").ColumnWidth = 18
   
   End With
   
   g_rst_Princi.MoveFirst
     
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
               
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!DATGEN_TITULO)
      
      If g_rst_Princi!DATGEN_PRYMCS = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = "PROYECTO VINCULADO"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = "PROYECTO NO VINCULADO"
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!PRODUC_DESCRI) 'moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!CLIENTE) 'moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!INSTANCIA) 'moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SOLMAE_CODINS))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!SITUACION) 'moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SOLMAE_SITINS))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!SOLMAE_CONHIP)
              
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = "SOLES"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = "DOLARES AMERICANOS"
      End If
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = 0
      End If
         
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = 0
      End If
              
      If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Or g_rst_Princi!SOLMAE_COMVTA_DOL > 0 Then
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_SOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00") * 100) + "%"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_DOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00") * 100) + "%"
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = CStr(0) + "%"
      End If
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
      End If
         
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = 0
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
Private Sub fs_Limpia()
   'cmb_TipPry.Clear
   chk_TipPry.Value = 0
   'cmb_Proyec.Clear
   chk_Proyec.Value = 0
   'cmb_TipCon.Clear
   chk_TipCon.Value = 0
End Sub
