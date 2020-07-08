VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.MDIForm frm_MnuPri_01 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8235
   ClientLeft      =   1485
   ClientTop       =   1950
   ClientWidth     =   11955
   Icon            =   "AteCli_frm_525.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   14430
      Top             =   9090
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11955
      _Version        =   65536
      _ExtentX        =   21087
      _ExtentY        =   1138
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
      Begin VB.CommandButton cmd_Seguim 
         Height          =   585
         Left            =   2430
         Picture         =   "AteCli_frm_525.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Seguimiento de Solicitud de Crédito Hipotecario"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_TipCam 
         Height          =   585
         Left            =   1830
         Picture         =   "AteCli_frm_525.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Consulta Tipo de Cambio"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_ConCre 
         Height          =   585
         Left            =   3630
         Picture         =   "AteCli_frm_525.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Consulta de Crédito Hipotecario"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_ConSol 
         Height          =   585
         Left            =   3030
         Picture         =   "AteCli_frm_525.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Consulta de Solicitud de Crédito Hipotecario"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_SimCre 
         Height          =   585
         Left            =   1230
         Picture         =   "AteCli_frm_525.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Simulación de Créditos Hipotecarios"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_CamCon 
         Height          =   585
         Left            =   630
         Picture         =   "AteCli_frm_525.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cambio de Contraseña"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   30
         Picture         =   "AteCli_frm_525.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir de Plataforma"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   7
      Top             =   7845
      Width           =   11955
      _Version        =   65536
      _ExtentX        =   21087
      _ExtentY        =   688
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
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Begin Threed.SSPanel pnl_EntDat 
         Height          =   315
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   3900
         _Version        =   65536
         _ExtentX        =   6879
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "lm_db_db1 - prod1"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
      End
      Begin Threed.SSPanel pnl_NumVer 
         Height          =   315
         Left            =   3960
         TabIndex        =   9
         Top             =   30
         Width           =   2100
         _Version        =   65536
         _ExtentX        =   3704
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "rev. 008-1028.1"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
      End
      Begin Threed.SSPanel pnl_TipCam 
         Height          =   315
         Left            =   6090
         TabIndex        =   10
         Top             =   30
         Width           =   5400
         _Version        =   65536
         _ExtentX        =   9525
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Tipo Cambio: Compra: S/. 2.00 - Venta: S/. 2.01"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
      End
   End
   Begin VB.Menu mnuAte 
      Caption         =   "Atención Comercial"
      Begin VB.Menu mnuAte_Opcion 
         Caption         =   "Simulación de Créditos Hipotecarios"
         Index           =   1
      End
      Begin VB.Menu mnuAte_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuAte_Opcion 
         Caption         =   "Registro de Clientes Potenciales"
         Index           =   3
      End
      Begin VB.Menu mnuAte_Opcion 
         Caption         =   "Registro de Prospectos"
         Index           =   4
      End
      Begin VB.Menu mnuAte_Opcion 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuAte_Opcion 
         Caption         =   "Ingreso de Solicitud de Crédito Hipotecario"
         Index           =   6
      End
      Begin VB.Menu mnuAte_Opcion 
         Caption         =   "Envío de Solicitudes a Créditos"
         Index           =   7
      End
      Begin VB.Menu mnuAte_Opcion 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuAte_Opcion 
         Caption         =   "Seguimiento de Solicitudes de Créditos Hipotecarios"
         Index           =   9
      End
      Begin VB.Menu mnuAte_Opcion 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuAte_Opcion 
         Caption         =   "Registro de Planes de Ahorro"
         Index           =   11
      End
   End
   Begin VB.Menu mnuAdm 
      Caption         =   "Administración"
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Gestión de Proyectos No Vinculados"
         Index           =   1
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Asignacion de Consejeros a Proyectos"
         Index           =   2
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Cambio de Consejero Hipotecario"
         Index           =   4
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Cambio de Ejecutivo de Seguimiento"
         Index           =   5
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Anulación de Solicitud de Crédito Hipotecario"
         Index           =   7
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Actualización Tasa de Solicitud de Crédito Hipotecario"
         Index           =   8
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Gestión de Clientes"
         Index           =   10
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Actualización Datos Cliente"
         Index           =   11
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Simulación de Prepago Parcial"
         Index           =   13
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Simulación de Prepago Total"
         Index           =   14
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Cambio de Fecha de Pago"
         Index           =   15
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuAdm_Opcion 
         Caption         =   "Emisión Conformidad AFP"
         Index           =   17
      End
   End
   Begin VB.Menu mnuCon 
      Caption         =   "Consultas"
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Solicitud de Crédito Hipotecario"
         Index           =   1
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Crédito Hipotecario"
         Index           =   2
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Tipo de Cambio"
         Index           =   4
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Posición de Solicitudes en Trámite por Consejero Hipotecario"
         Index           =   6
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Posición de Solicitudes en Trámite"
         Index           =   7
      End
   End
   Begin VB.Menu mnuRpt 
      Caption         =   "Reportes"
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes en Trámite"
         Index           =   1
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes en Trámite con Apr. Crediticia y Pago de Gastos"
         Index           =   3
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes en Trámite con Apr. Crediticia sin Pago de Gastos"
         Index           =   5
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes en Trámite por Proyecto "
         Index           =   7
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes en Trámite Observadas"
         Index           =   9
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes en Trámite por Instancia"
         Index           =   11
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes en Trámite con Apr. Condicionada"
         Index           =   13
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes Rechazadas"
         Index           =   15
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes Desembolsadas"
         Index           =   17
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes Desembolsadas (Proyecto Inmobiliario)"
         Index           =   18
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Resumen de Solicitudes Desembolsadas"
         Index           =   19
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Cuadro de Seguimiento de Solicitudes"
         Index           =   21
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Cuadro de Seguimiento de Solicitudes por Fechas"
         Index           =   22
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   23
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes por Fecha de Ingreso"
         Index           =   24
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Seguimiento de Solicitudes por Tiempos en Instancias"
         Index           =   25
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   26
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Seguimiento de Prospectos"
         Index           =   27
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Desembolsos Mensual"
         Index           =   28
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Tuberia"
         Index           =   29
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Seguimiento de Proyectos"
         Index           =   30
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Desembolsos Acumulado"
         Index           =   31
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte Inspektor"
         Index           =   32
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
   If mnuCon_Opcion(2).Enabled Then
      Call mnuCon_Opcion_Click(2)
   Else
      MsgBox "Usuario no tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_ConSol_Click()
   If mnuCon_Opcion(1).Enabled Then
      Call mnuCon_Opcion_Click(1)
   Else
      MsgBox "Usuario no tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   End If
End Sub

Private Sub cmd_Seguim_Click()
   If mnuAte_Opcion(5).Enabled Then
      frm_Seg_SolHip_52.Show 1
   Else
      MsgBox "Usuario no tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_SimCre_Click()
   If mnuAte_Opcion(1).Enabled Then
      Call mnuAte_Opcion_Click(1)
   Else
      MsgBox "Usuario no tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_TipCam_Click()
   If mnuCon_Opcion(4).Enabled Then
      Call mnuCon_Opcion_Click(4)
   Else
      MsgBox "Usuario no tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub MDIForm_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   'Call fs_HabSeg
   Call moddat_gf_Cargar_AgrPrd
   
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

Private Sub mnuAdm_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Proyectos No Vinculados
         frm_PryNVi_01.Show 1
      
      Case 2
         'Asignar Proyectos a Consejeros Hipotecarios
         frm_PryAsig_01.Show 1
         
      Case 4
         'Cambio de Consejero Hipotecario
         moddat_g_int_FlgPre = 5
         frm_CamCon_11.Show 1
         
      Case 5
         'Cambio de Ejecutivo de Seguimiento
         frm_CamSeg_01.Show 1
         
      Case 7
         'Anulación de Solicitudes
         frm_AnuSol_01.Show 1
         
      Case 8
         'Modificación de Solicitudes
         moddat_g_int_FlgPre = 6
         frm_CamCon_11.Show 1
         
      Case 10
         'Mantenimiento de Clientes
         moddat_g_int_FlgCre = 1
         frm_MntCli_51.Show 1
         
      Case 11
         'Actualización Datos del Clientes
         moddat_g_int_FlgCre = 2
         frm_MntCli_51.Show 1
         
      Case 13
         'Simulacion de Prepago Parcial
         moddat_g_int_FlgPre = 2
         frm_Con_CreHip_51.Show 1
   
      Case 14
         'Simulacion de Prepago Total
         moddat_g_int_FlgPre = 3
         frm_Con_CreHip_51.Show 1
   
      Case 15
         'Cambio de Fecha de Pago
         moddat_g_int_FlgPre = 4
         frm_Con_CreHip_51.Show 1
   
      Case 17
         frm_Tra_CarAFP_01.Show 1
   End Select
End Sub

Private Sub mnuAte_Opcion_Click(Index As Integer)


   Select Case Index
      Case 1
         'Simulación de Créditos Hipotecarios
         If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
            MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         'Call abc
         frm_SimCre_11.Show 1
      
      Case 3
         'Ingreso de Prospecto de Clientes Potenciales Crédito Hipotecario
         'MsgBox ("xx")
         'Call abc
         frm_IngCliPot_01.Show 1
         
      Case 4
         'Ingreso de Prospecto de Crédito Hipotecario
         frm_IngPros_01.Show 1
         
      Case 6
         'Ingreso de Solicitud de Crédito Hipotecario
         If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
            MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         frm_SolCre_51.Show 1
         
      Case 7
         frm_SolCre_56.Show 1
         
      Case 9
         'Seguimiento de Solicitud de Crédito Hipotecario
         frm_Seg_SolHip_51.Show 1
         
      Case 11
         'Registro del plan de ahorros
        ' Call abc
         frm_Pla_Aho_01.Show 1
   End Select
End Sub

Private Sub mnuCon_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Consulta de Solicitud de Crédito Hipotecario
         frm_Con_SolHip_51.Show 1
         
      Case 2
         'Consulta de Crédito Hipotecario
         moddat_g_int_FlgPre = 1
         frm_Con_CreHip_51.Show 1
         
      Case 4
         'Consulta de Tipo de Cambio
         frm_ConTCa_01.Show 1
         
      Case 6
         'Posición x Consejero Hipotecario (Trámite)
         frm_ActCon_01.Show 1
   
      Case 7
         'Posición General (Trámite)
         frm_ActCon_02.Show 1
   End Select
End Sub

Private Sub mnuRpt_Opcion_Click(Index As Integer)

Select Case Index
      Case 1
         'Reporte de Solicitudes en Trámite
         'Reemplaza los formularios frm_RptSol_08 y  frm_RptSol_09
         frm_RptSol_42.Show 1

      Case 3
         'Reporte de Solicitudes en Trámite con Apr. Crediticia y con Pago de Gastos
         'Reemplaza los formularios frm_RptSol_10 y  frm_RptSol_11
         frm_RptSol_43.Show 1

      Case 5
         'Reporte de Solicitudes en Trámite con Apr. Crediticia y sin Pago de Gastos
         'Reemplaza los formularios frm_RptSol_12 y  frm_RptSol_13
         frm_RptSol_44.Show 1

      Case 7
         'Reporte de Solicitudes en Tramite x Proyecto
         'Reemplaza los formularios frm_RptSol_16 y  frm_RptSol_17
         frm_RptSol_45.Show 1

      Case 9
         'Reporte de Solicitudes en Trámite Observadas
         'Reemplaza los formularios frm_RptSol_18 y  frm_RptSol_19
         frm_RptSol_46.Show 1

      Case 11
         'Reporte de Solicitudes en Trámite x Instancias
         'Reemplaza los formularios frm_RptSol_15, 24, 25, 26, 27 y frm_RptSol_29
         frm_RptSol_47.Show 1
     
      Case 13
         'Reporte de Solicitudes en Trámite con Aprobación Condicionada
         frm_RptSol_07.Show 1

      Case 15
         'Reporte de Solicitudes Rechazadas
         'Reemplaza los formularios frm_RptSol_22 y  frm_RptSol_23
         frm_RptSol_48.Show 1

      Case 17
         'Reporte de Solicitudes Desembolsadas
         'Reemplaza los formularios frm_RptSol_20 y  frm_RptSol_21
         frm_RptSol_49.Show 1

      Case 18
         'Reporte de Desembolsos (Proyecto Inmobiliario)
         frm_RptSol_14.Show 1

      Case 19
         'Resumen de Desembolsos x Consejero Hipotecario
         frm_RptSol_30.Show 1
      
      Case 21
         'Cuadro de Seguimiento de Solicitudes
         'Reemplaza los formularios frm_RptSol_28 y  frm_RptSol_31
         frm_RptSol_51.Show 1
      
      Case 22
         'Cuadro de Seguimiento de Solicitudes por Fechas
         frm_RptSol_34.Show 1
         
      Case 24
         'Reporte de Solicitudes por Fecha de Ingreso
         'Reemplaza los formularios frm_RptSol_32 y  frm_RptSol_33
         frm_RptSol_50.Show 1
         
      Case 25
         'Cuadro de Seguimiento de Solicitudes por incidencias
         frm_RptSol_35.Show 1
   
      Case 27
         'Reporte de Seguimiento de Prospectos
         frm_RptSol_36.Show 1
         
      Case 28
         'Reporte de Desembolsos por Promotor
         frm_RptSol_37.Show 1
      
      Case 29
         'Reporte de Tuberia
         frm_RptSol_38.Show 1
      
      Case 30
         'Reporte de Seguimiento de Proyectos
         frm_RptSol_39.Show 1
         
      Case 31
         'Reporte Acumulado Mensual de Desembolsos
         frm_RptSol_40.Show 1
         
       Case 32
         'Reporte Inspektor
        ' MsgBox ("clic aki")
         frm_RpIpk_01.Show 1
         
   End Select

End Sub

Private Sub fs_HabSeg()
Dim r_int_Posici     As Integer
Dim r_str_CodMen     As String
Dim r_dbl_TipVta     As Double
Dim r_dbl_TipCom     As Double
   
   'pnl_Seg_NomUsu.Caption = modgen_g_str_CodUsu
   pnl_NumVer.Caption = modgen_g_str_NumRev
   pnl_EntDat.Caption = moddat_g_str_NomEsq & " - " & UCase(moddat_g_str_EntDat)
   r_dbl_TipVta = moddat_gf_ObtieneTipCamDia(1, 2, Format(date, "yyyymmdd"), 1)
   r_dbl_TipCom = moddat_gf_ObtieneTipCamDia(1, 2, Format(date, "yyyymmdd"), 2)
   pnl_TipCam.Caption = "Tipo de Cambio: Compra: S/. " & Format(r_dbl_TipCom, "###0.0000") & " - Venta: S/. " & Format(r_dbl_TipVta, "###0.0000")
   
   'Desactivando todas las opciones
   For r_int_Posici = 1 To mnuAte_Opcion.Count
      If mnuAte_Opcion(r_int_Posici).Caption <> "-" Then
         mnuAte_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici

   For r_int_Posici = 1 To mnuAdm_Opcion.Count
      If mnuAdm_Opcion(r_int_Posici).Caption <> "-" Then
         mnuAdm_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici

   For r_int_Posici = 1 To mnuCon_Opcion.Count
      If mnuCon_Opcion(r_int_Posici).Caption <> "-" Then
         mnuCon_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuRpt_Opcion.Count
      If mnuRpt_Opcion(r_int_Posici).Caption <> "-" Then
         mnuRpt_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'Verificando si todas las Opciones están habilitadas
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM SEG_PLTOPC "
   g_str_Parame = g_str_Parame & " WHERE PLTOPC_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "   AND PLTOPC_FLGMEN = 2 "
   g_str_Parame = g_str_Parame & "ORDER BY PLTOPC_CODMEN ASC, PLTOPC_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTOPC_CODMEN)
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTOPC_CODMEN)
            Select Case r_str_CodMen
               Case "MNUATE": mnuAte_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUADM": mnuAdm_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNURPT": mnuRpt_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Verificando por Plantilla de Acceso
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM SEG_PLTPLA "
   g_str_Parame = g_str_Parame & " WHERE PLTPLA_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "   AND PLTPLA_TIPUSU = '" & CStr(modgen_g_int_TipUsu) & "' "
   g_str_Parame = g_str_Parame & "ORDER BY PLTPLA_CODMEN ASC, PLTPLA_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTPLA_CODMEN)
         
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTPLA_CODMEN)
            Select Case r_str_CodMen
               Case "MNUATE": mnuAte_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUADM": mnuAdm_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNURPT": mnuRpt_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Verificando por Personalización de Opciones
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM SEG_PLTUSU "
   g_str_Parame = g_str_Parame & " WHERE PLTUSU_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "   AND PLTUSU_CODUSU = '" & CStr(modgen_g_str_CodUsu) & "' "
   g_str_Parame = g_str_Parame & "ORDER BY PLTUSU_CODMEN ASC, PLTUSU_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
         
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
            Select Case r_str_CodMen
               Case "MNUATE": mnuAte_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUADM": mnuAdm_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNURPT": mnuRpt_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

