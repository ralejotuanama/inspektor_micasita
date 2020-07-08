VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.MDIForm frm_MnuPri_01 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8235
   ClientLeft      =   2580
   ClientTop       =   2040
   ClientWidth     =   10380
   Icon            =   "AteCli_frm_056.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   14430
      Top             =   9090
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7860
      Width           =   10380
      _Version        =   65536
      _ExtentX        =   18309
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Begin Threed.SSPanel pnl_Seg_TipUsu 
         Height          =   315
         Left            =   8760
         TabIndex        =   6
         Top             =   30
         Width           =   6555
         _Version        =   65536
         _ExtentX        =   11562
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Tipo de Usuario: ADMINISTRADOR DE PLATAFORMA"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
         Alignment       =   1
      End
      Begin Threed.SSPanel pnl_Seg_NomUsu 
         Height          =   315
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   8685
         _Version        =   65536
         _ExtentX        =   15319
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Nombre Usuario: MIGUEL ANGEL IKEHARA PUNK"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
         Alignment       =   1
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10380
      _Version        =   65536
      _ExtentX        =   18309
      _ExtentY        =   1296
      _StockProps     =   15
      BackColor       =   -2147483633
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
      Begin VB.CommandButton cmd_TipCam 
         Height          =   675
         Left            =   2790
         Picture         =   "AteCli_frm_056.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Consulta Tipo de Cambio"
         Top             =   30
         Width           =   675
      End
      Begin VB.CommandButton cmd_ConCre 
         Height          =   675
         Left            =   1410
         Picture         =   "AteCli_frm_056.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Consulta de Crédito Hipotecario"
         Top             =   30
         Width           =   675
      End
      Begin VB.CommandButton cmd_ConSol 
         Height          =   675
         Left            =   720
         Picture         =   "AteCli_frm_056.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Consulta de Solicitud de Crédito Hipotecario"
         Top             =   30
         Width           =   675
      End
      Begin VB.CommandButton cmd_SimCre 
         Height          =   675
         Left            =   30
         Picture         =   "AteCli_frm_056.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Simulación de Créditos Hipotecarios"
         Top             =   30
         Width           =   675
      End
      Begin VB.CommandButton cmd_Calcul 
         Height          =   675
         Left            =   2100
         Picture         =   "AteCli_frm_056.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Calculadora"
         Top             =   30
         Width           =   675
      End
      Begin VB.CommandButton cmd_CamCon 
         Height          =   675
         Left            =   3480
         Picture         =   "AteCli_frm_056.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cambio de Contraseña"
         Top             =   30
         Width           =   675
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   675
         Left            =   4140
         Picture         =   "AteCli_frm_056.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir de Plataforma"
         Top             =   30
         Width           =   675
      End
   End
   Begin VB.Menu mnuVta 
      Caption         =   "&Ventas"
      Begin VB.Menu mnuVta_SimCre 
         Caption         =   "Simulación de Créditos"
      End
      Begin VB.Menu mnuVta_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVta_SolMCa 
         Caption         =   "Solicitud de Crédito Hipotecario miCasita"
      End
      Begin VB.Menu mnuVta_SolMCP 
         Caption         =   "Solicitud de Crédito Hipotecario miCasita - PBP"
      End
      Begin VB.Menu mnuVta_Line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVta_SolMV1 
         Caption         =   "Solicitud de Crédito Hipotecario Mivivienda - CRC-PBP"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVta_SolMV2 
         Caption         =   "Solicitud de Crédito Hipotecario Mivivienda - CME"
      End
      Begin VB.Menu mnuVta_SolMV3 
         Caption         =   "Solicitud de Crédito Hipotecario Mivivienda - Mihogar"
      End
      Begin VB.Menu mnuVta_Line03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVta_SegSol 
         Caption         =   "Seguimiento de Solicitud de Crédito Hipotecario"
      End
   End
   Begin VB.Menu mnuAdm 
      Caption         =   "Administración"
      Begin VB.Menu mnuAdm_PryNvi 
         Caption         =   "Gestión de Proyectos No Vinculados"
      End
      Begin VB.Menu mnuAdm_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdm_CamCon 
         Caption         =   "Cambio de Consejero Hipotecario"
      End
      Begin VB.Menu mnuAdm_CamSeg 
         Caption         =   "Cambio de Ejecutivo de Seguimiento"
      End
      Begin VB.Menu mnuAdm_Line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdm_AnuSol 
         Caption         =   "Anulación de Solicitud de Crédito Hipotecario"
      End
   End
   Begin VB.Menu mnuCon 
      Caption         =   "Consultas"
      Begin VB.Menu mnuCon_SolCre 
         Caption         =   "Solicitud de Crédito Hipotecario"
      End
      Begin VB.Menu mnuCon_CreHip 
         Caption         =   "Crédito Hipotecario"
      End
      Begin VB.Menu mnuCon_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCon_TipCam 
         Caption         =   "Tipo de Cambio"
      End
   End
   Begin VB.Menu mnuRep 
      Caption         =   "Reportes"
      Begin VB.Menu mnuRep_RepGen 
         Caption         =   "Reporte General de Solicitudes de Crédito Hipotecario"
      End
      Begin VB.Menu mnuRep_SegSol 
         Caption         =   "Reporte de Seguimiento de Solicitudes de Crédito Hipotecario"
      End
      Begin VB.Menu mnuRep_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRep_SolObs 
         Caption         =   "Reporte de Solicitudes Observadas"
      End
      Begin VB.Menu mnuRep_Line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRep_PenEva 
         Caption         =   "Reporte de Solicitudes Pendientes de Evaluación Crediticia"
      End
      Begin VB.Menu mnuRep_EvaCre 
         Caption         =   "Reporte de Solicitudes en Evaluación Crediticia"
      End
      Begin VB.Menu mnuRep_AprCre 
         Caption         =   "Reporte de Solicitudes Aprobadas Crediticiamente"
      End
      Begin VB.Menu mnuRep_Line03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRep_EstCon 
         Caption         =   "Estadística de Solicitudes por Consejero Hipotecario"
      End
   End
End
Attribute VB_Name = "frm_MnuPri_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Calcul_Click()
   Dim r_lng_NumPid    As Long
   
   r_lng_NumPid = Shell("c:\windows\system32\calc.exe", vbNormalFocus)
   
   If r_lng_NumPid = 0 Then
      MsgBox "Error Iniciando la Aplicación", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_CamCon_Click()
   If modgen_g_str_CodUsu <> "DESARROLLO" Then
      frm_IdeUsu_02.Show 1
   End If
End Sub

Private Sub cmd_ConCre_Click()
   Call mnuCon_CreHip_Click
End Sub

Private Sub cmd_ConSol_Click()
   Call mnuCon_SolCre_Click
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   End If
End Sub

Private Sub cmd_SimCre_Click()
   Call mnuAte_SimCre_Click
End Sub

Private Sub cmd_TipCam_Click()
   frm_ConTCa_01.Show 1
End Sub

Private Sub MDIForm_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt & " - " & modgen_g_str_NumRev & " [" & moddat_g_str_NomEsq & " - " & moddat_g_str_EntDat & "]"

   pnl_Seg_NomUsu.Caption = "Nombre de Usuario: " & modgen_g_str_NomUsu
   pnl_Seg_TipUsu.Caption = "Tipo de Usuario: " & Mid(moddat_gf_Consulta_ParDes("351", CStr(modgen_g_int_TipUsu)), 10)
   
   Call fs_Activa_MenPri
   
   Screen.MousePointer = 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   Else
      Cancel = True
   End If
End Sub

Private Sub mnuAdm_CamCon_Click()
   frm_CamCon_11.Show 1
End Sub

Private Sub mnuAdm_PryNvi_Click()
   frm_PryNVi_01.Show 1
End Sub

Private Sub mnuAte_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_02.Show 1
End Sub

Private Sub mnuCon_CreHip_Click()
   frm_ConCre_51.Show 1
End Sub

Private Sub mnuCon_SolCre_Click()
   frm_ConSol_49.Show 1
End Sub

Private Sub mnuCon_TipCam_Click()
   frm_ConTCa_01.Show 1
End Sub

Private Sub mnuHer_ConTCa_Click()
   Call cmd_TipCam_Click
End Sub

Private Sub mnuRep_AprCre_Click()
   frm_RptSol_02.Show 1
End Sub

Private Sub mnuRep_EstCon_Click()
   frm_EstSol_01.Show 1
End Sub

Private Sub mnuRep_EvaCre_Click()
   frm_RptSol_05.Show 1
End Sub

Private Sub mnuRep_PenEva_Click()
   frm_RptSol_06.Show 1
End Sub

Private Sub mnuRep_RepGen_Click()
   frm_RptSol_01.Show 1
End Sub

Private Sub mnuRep_SegSol_Click()
   frm_RptSol_03.Show 1
End Sub

Private Sub mnuRep_SolObs_Click()
   frm_RptSol_04.Show 1
End Sub

Private Sub mnuAdm_AnuSol_Click()
   frm_AnuSol_01.Show 1
End Sub

Private Sub mnuAdm_CamSeg_Click()
   frm_CamSeg_01.Show 1
End Sub

Private Sub mnuVta_SegSol_Click()
   frm_SegSol_15.Show 1
End Sub

Private Sub mnuVta_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_02.Show 1
End Sub

Private Sub mnuVta_SolMCa_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_CodPrd = "002"
   frm_SolCre_01.Show 1
End Sub

Private Sub fs_Activa_MenPri()
   'Desactivando todas las opciones
   mnuVta_SimCre.Enabled = False
   mnuVta_SolMCa.Enabled = False
   mnuVta_SolMCP.Enabled = False
   mnuVta_SolMV2.Enabled = False
   mnuVta_SolMV3.Enabled = False
   mnuVta_SegSol.Enabled = False
   
   mnuAdm_PryNvi.Enabled = False
   mnuAdm_CamCon.Enabled = False
   mnuAdm_CamSeg.Enabled = False
   mnuAdm_AnuSol.Enabled = False
   
   mnuCon_SolCre.Enabled = False
   mnuCon_CreHip.Enabled = False
   mnuCon_TipCam.Enabled = False
   
   mnuRep.Enabled = False
   
   Select Case modgen_g_int_TipUsu
      Case 1000
         mnuVta_SimCre.Enabled = True
         mnuVta_SolMCa.Enabled = True
         mnuVta_SolMCP.Enabled = True
         mnuVta_SolMV2.Enabled = True
         mnuVta_SolMV3.Enabled = True
         mnuVta_SegSol.Enabled = True
         
         mnuAdm_PryNvi.Enabled = True
         mnuAdm_CamCon.Enabled = True
         mnuAdm_CamSeg.Enabled = True
         mnuAdm_AnuSol.Enabled = True
         
         mnuCon_SolCre.Enabled = True
         mnuCon_CreHip.Enabled = True
         mnuCon_TipCam.Enabled = True
         
         mnuRep.Enabled = True
      
      Case 20100      'Director Comercial
         mnuVta_SimCre.Enabled = True
         mnuVta_SolMCa.Enabled = True
         mnuVta_SolMV2.Enabled = True
         mnuVta_SolMV3.Enabled = True
         mnuVta_SegSol.Enabled = True
         
         mnuAdm_PryNvi.Enabled = True
         mnuAdm_CamCon.Enabled = True
         mnuAdm_CamSeg.Enabled = True
         mnuAdm_AnuSol.Enabled = True
         
         mnuCon_SolCre.Enabled = True
         mnuCon_CreHip.Enabled = True
         mnuCon_TipCam.Enabled = True
         
         mnuRep.Enabled = True
      
      Case 20110      'Jefe de Seguimiento
         mnuVta_SimCre.Enabled = True
         mnuVta_SolMCa.Enabled = True
         mnuVta_SolMV2.Enabled = True
         mnuVta_SolMV3.Enabled = True
         mnuVta_SolMCP.Enabled = True
         mnuVta_SegSol.Enabled = True
         
         mnuAdm_PryNvi.Enabled = True
         mnuAdm_CamCon.Enabled = True
         mnuAdm_CamSeg.Enabled = True
         mnuAdm_AnuSol.Enabled = True
         
         mnuCon_SolCre.Enabled = True
         mnuCon_CreHip.Enabled = True
         mnuCon_TipCam.Enabled = True
         
         mnuRep.Enabled = True
      
      Case 20111      'Ejecutivo de Seguimiento
         mnuVta_SimCre.Enabled = True
         mnuVta_SolMCa.Enabled = True
         mnuVta_SolMCP.Enabled = True
         mnuVta_SolMV2.Enabled = True
         mnuVta_SolMV3.Enabled = True
         mnuVta_SegSol.Enabled = True
         
         mnuCon_SolCre.Enabled = True
         mnuCon_CreHip.Enabled = True
         mnuCon_TipCam.Enabled = True
         
      Case 20120      'Jefe de Ventas
         mnuVta_SimCre.Enabled = True
         mnuVta_SolMCa.Enabled = True
         mnuVta_SolMV2.Enabled = True
         mnuVta_SolMV3.Enabled = True
         mnuVta_SegSol.Enabled = True
         
         mnuCon_SolCre.Enabled = True
         mnuCon_CreHip.Enabled = True
         mnuCon_TipCam.Enabled = True
      
      Case 20121      'Consejero Hipotecario
         mnuVta_SimCre.Enabled = True
         mnuVta_SolMCa.Enabled = True
         mnuVta_SolMV2.Enabled = True
         mnuVta_SolMV3.Enabled = True
         mnuVta_SegSol.Enabled = True
         
         mnuCon_SolCre.Enabled = True
         mnuCon_CreHip.Enabled = True
         mnuCon_TipCam.Enabled = True
      
      Case 20200      'Director de Producción
         mnuVta_SimCre.Enabled = True
         mnuVta_SolMCa.Enabled = True
         mnuVta_SolMV2.Enabled = True
         mnuVta_SolMV3.Enabled = True
         mnuVta_SegSol.Enabled = True
         
         mnuCon_SolCre.Enabled = True
         mnuCon_CreHip.Enabled = True
         mnuCon_TipCam.Enabled = True
      
      Case 20900      'Consulta
         mnuVta_SimCre.Enabled = True
      
         mnuCon_SolCre.Enabled = True
         mnuCon_CreHip.Enabled = True
         mnuCon_TipCam.Enabled = True
      
   End Select
End Sub

Private Sub mnuVta_SolMCP_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_CodPrd = "006"
   frm_SolCre_01.Show 1
End Sub

Private Sub mnuVta_SolMV1_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_CodPrd = "001"
   frm_SolCre_01.Show 1
End Sub

Private Sub mnuVta_SolMV2_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_CodPrd = "003"
   
   frm_SolCre_01.Show 1
End Sub

Private Sub mnuVta_SolMV3_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_CodPrd = "004"
   
   frm_SolCre_01.Show 1
End Sub
