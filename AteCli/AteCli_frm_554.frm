VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Pla_Aho_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9195
   ClientLeft      =   1725
   ClientTop       =   3720
   ClientWidth     =   14340
   Icon            =   "AteCli_frm_554.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   14340
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9195
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   14370
      _Version        =   65536
      _ExtentX        =   25347
      _ExtentY        =   16219
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
         TabIndex        =   9
         Top             =   60
         Width           =   14250
         _Version        =   65536
         _ExtentX        =   25135
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
            Left            =   630
            TabIndex        =   10
            Top             =   60
            Width           =   8490
            _Version        =   65536
            _ExtentX        =   14975
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Registro de Planes de Ahorro"
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
            Picture         =   "AteCli_frm_554.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   11
         Top             =   780
         Width           =   14250
         _Version        =   65536
         _ExtentX        =   25135
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
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   180
            Width           =   2985
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13650
            Picture         =   "AteCli_frm_554.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   45
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   645
            Picture         =   "AteCli_frm_554.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Modificar Plan de Ahorro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   45
            Picture         =   "AteCli_frm_554.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Adicionar Plan de Ahorro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   585
            Left            =   1245
            Picture         =   "AteCli_frm_554.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Anular Plan de Ahorro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consulta 
            Height          =   585
            Left            =   1830
            Picture         =   "AteCli_frm_554.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Consultar registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_PlnCuo 
            Height          =   585
            Left            =   2430
            Picture         =   "AteCli_frm_554.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Consultar Cuotas"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   3030
            Picture         =   "AteCli_frm_554.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Situación:"
            Height          =   195
            Left            =   5520
            TabIndex        =   23
            Top             =   240
            Width           =   705
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   7665
         Left            =   60
         TabIndex        =   12
         Top             =   1470
         Width           =   14250
         _Version        =   65536
         _ExtentX        =   25135
         _ExtentY        =   13520
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   7245
            Left            =   60
            TabIndex        =   0
            Top             =   360
            Width           =   14145
            _ExtentX        =   24950
            _ExtentY        =   12779
            _Version        =   393216
            Rows            =   26
            Cols            =   19
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   4815
            TabIndex        =   13
            Top             =   60
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1905
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "DOI Cliente"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_PriVct 
            Height          =   285
            Left            =   8805
            TabIndex        =   14
            Top             =   60
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Registro"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   5895
            TabIndex        =   15
            Top             =   60
            Width           =   2910
            _Version        =   65536
            _ExtentX        =   5133
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   1620
            TabIndex        =   16
            Top             =   60
            Width           =   3210
            _Version        =   65536
            _ExtentX        =   5662
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   90
            TabIndex        =   17
            Top             =   60
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "N° Operación"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_NumMes 
            Height          =   285
            Left            =   9840
            TabIndex        =   18
            Top             =   60
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Meses"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_MtoAho 
            Height          =   285
            Left            =   10440
            TabIndex        =   19
            Top             =   60
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1905
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto. Ahorro"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_Situac 
            Height          =   285
            Left            =   11520
            TabIndex        =   20
            Top             =   60
            Width           =   1110
            _Version        =   65536
            _ExtentX        =   1958
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_Consejero 
            Height          =   285
            Left            =   12630
            TabIndex        =   21
            Top             =   60
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Consejero"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frm_Pla_Aho_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Agrega_Click()
   modmip_g_int_FlgGrb_1 = 1
   modmip_g_int_FlgAct_1 = 2
   frm_Pla_Aho_02.Show 1
      
   Screen.MousePointer = 11
   Call fs_Carga_grid
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Editar_Click()
Dim r_int_CodAct     As Integer
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 9
   moddat_g_str_TipDoc = CInt(grd_Listad.Text)
   grd_Listad.Col = 10
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   grd_Listad.Col = 11
   moddat_g_str_NumOpe = CStr(grd_Listad.Text)
   grd_Listad.Col = 12
   r_int_CodAct = CInt(grd_Listad.Text)
   
   If r_int_CodAct = 8 Then
       MsgBox "Operacion esta ANULADA, no puede modificarse.", vbInformation, modgen_g_str_NomPlt
       Call gs_RefrescaGrid(grd_Listad)
       Exit Sub
   End If
   If r_int_CodAct = 9 Then
       MsgBox "Operacion esta CANCELADA, no puede modificarse.", vbInformation, modgen_g_str_NomPlt
       Call gs_RefrescaGrid(grd_Listad)
       Exit Sub
   End If
   
   Call gs_RefrescaGrid(grd_Listad)
   modmip_g_int_FlgGrb_1 = 2
   modmip_g_int_FlgAct_1 = 2
   
   frm_Pla_Aho_02.Show 1
   
   Screen.MousePointer = 11
   Call fs_Carga_grid
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Rechaz_Click()
Dim r_int_CodAct     As Integer
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 9
   moddat_g_str_TipDoc = CInt(grd_Listad.Text)
   grd_Listad.Col = 10
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   grd_Listad.Col = 11
   moddat_g_str_NumOpe = CStr(grd_Listad.Text)
   grd_Listad.Col = 12
   r_int_CodAct = CInt(grd_Listad.Text)

   If r_int_CodAct = 8 Then
       MsgBox "Operacion ya esta ANULADA, no puede anularse nuevamente.", vbInformation, modgen_g_str_NomPlt
       Call gs_RefrescaGrid(grd_Listad)
       Exit Sub
   End If
   If r_int_CodAct = 9 Then
       MsgBox "Operacion esta CANCELADA, no puede anularse.", vbInformation, modgen_g_str_NomPlt
       Call gs_RefrescaGrid(grd_Listad)
       Exit Sub
   End If

   MsgBox "Esta opción anula la operación: " & Mid(moddat_g_str_NumOpe, 1, 4) & "-" & Mid(moddat_g_str_NumOpe, 5, 8) & "-" & Mid(moddat_g_str_NumOpe, 13, 3), vbExclamation, modgen_g_str_NomPlt
   Call gs_RefrescaGrid(grd_Listad)
   DoEvents
   
   If MsgBox("¿Está seguro de anular la información del plan de ahorro del cliente?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'actualiza Información del Cliente en el plan de ahorro y cuotas
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CRE_AHOMAE "
   g_str_Parame = g_str_Parame & "   SET AHOMAE_SITUAC = '8' "
   g_str_Parame = g_str_Parame & " WHERE AHOMAE_NUMERO = '" & CStr(moddat_g_str_NumOpe) & "' "
   g_str_Parame = g_str_Parame & "   AND AHOMAE_TIPDOC = " & CStr(moddat_g_str_TipDoc) & " "
   g_str_Parame = g_str_Parame & "   AND AHOMAE_NUMDOC = '" & Trim(moddat_g_str_NumDoc) & "' "
   g_str_Parame = g_str_Parame & "   AND AHOMAE_SITUAC = '2' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       MsgBox "Error al ejecutar la eliminacion del registro de CRE_AHOMAE.", vbCritical, modgen_g_str_NomPlt
       Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CRE_AHOCUO "
   g_str_Parame = g_str_Parame & "   SET AHOCUO_SITUAC = 8 "
   g_str_Parame = g_str_Parame & " WHERE AHOCUO_NUMERO = '" & CStr(moddat_g_str_NumOpe) & "' "
   g_str_Parame = g_str_Parame & "   AND AHOCUO_SITUAC = 2 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       MsgBox "Error al ejecutar la eliminacion del registro de CRE_AHOCUO.", vbCritical, modgen_g_str_NomPlt
       Exit Sub
   End If

   Screen.MousePointer = 11
   Call fs_Carga_grid
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Consulta_Click()
Dim r_int_CodAct     As Integer
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 9
   moddat_g_str_TipDoc = CInt(grd_Listad.Text)
   grd_Listad.Col = 10
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   grd_Listad.Col = 11
   moddat_g_str_NumOpe = CStr(grd_Listad.Text)
   grd_Listad.Col = 12
   r_int_CodAct = CInt(grd_Listad.Text)

   Call gs_RefrescaGrid(grd_Listad)
   modmip_g_int_FlgGrb_1 = 3
   
   frm_Pla_Aho_02.Show 1
End Sub

Private Sub cmd_PlnCuo_Click()
    moddat_g_str_NumDoc = ""
    moddat_g_str_Codigo = ""
    moddat_g_str_Codigo = Trim(grd_Listad.TextMatrix(grd_Listad.RowSel, 11))
    moddat_g_str_NumDoc = Trim(grd_Listad.TextMatrix(grd_Listad.RowSel, 2))
    frm_Pla_Aho_03.Show 1
End Sub

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt

   Call fs_Inicio
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(grd_Listad)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   grd_Listad.Cols = 19 '15
   grd_Listad.ColWidth(0) = 1530
   grd_Listad.ColWidth(1) = 3200
   grd_Listad.ColWidth(2) = 1080
   grd_Listad.ColWidth(3) = 2900
   grd_Listad.ColWidth(4) = 1035
   grd_Listad.ColWidth(5) = 600
   grd_Listad.ColWidth(6) = 1100
   grd_Listad.ColWidth(7) = 1100
   grd_Listad.ColWidth(8) = 1250
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColWidth(13) = 0
   grd_Listad.ColWidth(14) = 0
   grd_Listad.ColWidth(15) = 0
   grd_Listad.ColWidth(16) = 0
   grd_Listad.ColWidth(17) = 0
   grd_Listad.ColWidth(18) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
   
   cmb_Situac.AddItem "VIGENTE"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 2
   cmb_Situac.AddItem "ANULADO"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 8
   cmb_Situac.AddItem "CANCELADO"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 9
   cmb_Situac.AddItem "TODOS"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 0
   cmb_Situac.ListIndex = 0
End Sub
Private Sub fs_Carga_grid()
Dim g_str_Parame    As String
 
   Call gs_LimpiaGrid(grd_Listad)
   moddat_g_str_Moneda = ""

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "    SELECT A.AHOMAE_NUMERO, A.AHOMAE_FECINI, A.AHOMAE_CONHIP, A.AHOMAE_PRIVCT, A.AHOMAE_NUMMES, "
   g_str_Parame = g_str_Parame & "           A.AHOMAE_MONAHO, A.AHOMAE_MTOAHO, A.AHOMAE_SITUAC, D.PRODUC_DESCRI, C.AHOCLI_TIPDOC, "
   g_str_Parame = g_str_Parame & "           C.AHOCLI_NUMDOC, C.AHOCLI_APEPAT, C.AHOCLI_APEMAT, C.AHOCLI_NOMBRE, A.AHOMAE_CUOPAG, "
   g_str_Parame = g_str_Parame & "           A.AHOMAE_CUOPEN, A.AHOMAE_CAPPAG, A.AHOMAE_CAPPEN "
   g_str_Parame = g_str_Parame & "      FROM CRE_AHOMAE A "
   g_str_Parame = g_str_Parame & "           INNER JOIN CRE_AHOCLI C ON A.AHOMAE_TIPDOC = C.AHOCLI_TIPDOC AND A.AHOMAE_NUMDOC = C.AHOCLI_NUMDOC "
   g_str_Parame = g_str_Parame & "           INNER JOIN CRE_PRODUC D ON D.PRODUC_CODIGO = A.AHOMAE_CODPRD "
   
   If modgen_g_int_TipUsu = 20121 Then   'Si Tipo de Usuario es Consejero Hipotecario
      g_str_Parame = g_str_Parame & " WHERE A.AHOMAE_CONHIP = '" & modgen_g_str_CodUsu & "' "
   Else
      g_str_Parame = g_str_Parame & " WHERE LENGTH(TRIM(A.AHOMAE_CONHIP)) > 0 "
   End If
   If (CInt(cmb_Situac.ItemData(cmb_Situac.ListIndex))) > 0 Then
      g_str_Parame = g_str_Parame & "   AND A.AHOMAE_SITUAC = " & CInt(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & " "
   End If
   g_str_Parame = g_str_Parame & " ORDER BY A.AHOMAE_NUMERO "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado clientes para esa selección.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Mid(Trim(g_rst_Princi!AHOMAE_NUMERO), 1, 4) & "-" & Mid(Trim(g_rst_Princi!AHOMAE_NUMERO), 5, 8) & "-" & Mid(Trim(g_rst_Princi!AHOMAE_NUMERO), 13, 3)
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!PRODUC_DESCRI & "")
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(g_rst_Princi!AHOCLI_NUMDOC & "")
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!AHOCLI_APEPAT & "") & " " & Trim(g_rst_Princi!AHOCLI_APEMAT & "") & " " & Trim(g_rst_Princi!AHOCLI_NOMBRE & "")
      grd_Listad.Col = 4
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!AHOMAE_FECINI)
      grd_Listad.Col = 5
      grd_Listad.Text = Format(Trim(g_rst_Princi!AHOMAE_NUMMES), "00")
      grd_Listad.Col = 6
      If g_rst_Princi!AHOMAE_MONAHO = 1 Then
         grd_Listad.Text = "S/.  " & Format(g_rst_Princi!AHOMAE_MTOAHO, "###,###,##0.00")
         moddat_g_str_Moneda = "SOLES"
      Else
         grd_Listad.Text = "US$  " & Format(g_rst_Princi!AHOMAE_MTOAHO, "###,###,##0.00")
         moddat_g_str_Moneda = "DÓLARES AMERICANOS"
      End If
      grd_Listad.Col = 7
      grd_Listad.Text = moddat_gf_Consulta_ParDes("027", g_rst_Princi!AHOMAE_SITUAC)
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(g_rst_Princi!AHOMAE_CONHIP)
      grd_Listad.Col = 9
      grd_Listad.Text = Trim(g_rst_Princi!AHOCLI_TIPDOC)
      grd_Listad.Col = 10
      grd_Listad.Text = Trim(g_rst_Princi!AHOCLI_NUMDOC)
      grd_Listad.Col = 11
      grd_Listad.Text = Trim(g_rst_Princi!AHOMAE_NUMERO)
      grd_Listad.Col = 12
      grd_Listad.Text = Trim(g_rst_Princi!AHOMAE_SITUAC)
      grd_Listad.Col = 13
      grd_Listad.Text = Trim(g_rst_Princi!AHOMAE_FECINI)
      grd_Listad.Col = 14
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!AHOMAE_PRIVCT)
      grd_Listad.Col = 15
      grd_Listad.Text = g_rst_Princi!AHOMAE_CUOPAG
      grd_Listad.Col = 16
      grd_Listad.Text = g_rst_Princi!AHOMAE_CAPPAG
'     If g_rst_Princi!AHOMAE_MONAHO = 1 Then
'        grd_Listad.Text = "S/.  " & Format(g_rst_Princi!AHOMAE_CAPPAG, "###,###,##0.00")
'     Else
'        grd_Listad.Text = "US$  " & Format(g_rst_Princi!AHOMAE_CAPPAG, "###,###,##0.00")
'     End If
      grd_Listad.Col = 17
      grd_Listad.Text = g_rst_Princi!AHOMAE_CUOPEN
      grd_Listad.Col = 18
      grd_Listad.Text = g_rst_Princi!AHOMAE_CAPPEN
'     If g_rst_Princi!AHOMAE_MONAHO = 1 Then
'        grd_Listad.Text = "S/.  " & Format(g_rst_Princi!AHOMAE_CAPPEN, "###,###,##0.00")
'     Else
'        grd_Listad.Text = "US$  " & Format(g_rst_Princi!AHOMAE_CAPPEN, "###,###,##0.00")
'     End If
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   Call gs_RefrescaGrid(grd_Listad)
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_DblClick()
    Call cmd_Editar_Click
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_nrofil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      r_int_nrofil = 1
      .Cells(1, 4) = "REPORTE DE PLANES DE AHORROS:"
      
      r_int_nrofil = r_int_nrofil + 2
      .Cells(r_int_nrofil, 1) = "OPERACION":             .Columns("A").ColumnWidth = 18
      .Cells(r_int_nrofil, 2) = "PRODUCTO":              .Columns("B").ColumnWidth = 40
      .Cells(r_int_nrofil, 3) = "DOI CLIENTE":           .Columns("C").ColumnWidth = 14
      .Cells(r_int_nrofil, 4) = "APELLIDOS Y NOMBRES":   .Columns("D").ColumnWidth = 40
      .Cells(r_int_nrofil, 5) = "F. REGISTRO":           .Columns("E").ColumnWidth = 16
      .Cells(r_int_nrofil, 6) = "PRIMER VCTO.":          .Columns("F").ColumnWidth = 16
      .Cells(r_int_nrofil, 7) = "MESES AHORRO":          .Columns("G").ColumnWidth = 16
      .Cells(r_int_nrofil, 8) = "MONTO AHORRO":          .Columns("H").ColumnWidth = 16
      .Cells(r_int_nrofil, 9) = "SITUACION":             .Columns("I").ColumnWidth = 16
      .Cells(r_int_nrofil, 10) = "CONSEJERO":            .Columns("J").ColumnWidth = 16
      .Cells(r_int_nrofil, 11) = "CUOTAS PAGAD.":       .Columns("K").ColumnWidth = 16
      .Cells(r_int_nrofil, 12) = "ABONO S/.":            .Columns("L").ColumnWidth = 16
      .Cells(r_int_nrofil, 13) = "CUOTAS PDTES.":        .Columns("M").ColumnWidth = 16
      .Cells(r_int_nrofil, 14) = "SALDO S/.":            .Columns("N").ColumnWidth = 16
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignLeft
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").HorizontalAlignment = xlHAlignRight
      .Columns("O").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 14)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 14)).HorizontalAlignment = xlHAlignCenter
      
      r_int_nrofil = r_int_nrofil + 1
      For r_int_nroaux = 0 To grd_Listad.Rows - 1
         .Cells(r_int_nrofil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_nrofil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_nrofil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_nrofil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_nrofil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_nrofil, 6) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 14)
         .Cells(r_int_nrofil, 7) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_nrofil, 8) = grd_Listad.TextMatrix(r_int_nroaux, 6)
         .Cells(r_int_nrofil, 9) = grd_Listad.TextMatrix(r_int_nroaux, 7)
         .Cells(r_int_nrofil, 10) = grd_Listad.TextMatrix(r_int_nroaux, 8)
         .Cells(r_int_nrofil, 11) = grd_Listad.TextMatrix(r_int_nroaux, 15)
         .Cells(r_int_nrofil, 12) = Format(grd_Listad.TextMatrix(r_int_nroaux, 16), "###,###,##0.00")
         .Cells(r_int_nrofil, 13) = grd_Listad.TextMatrix(r_int_nroaux, 17)
         .Cells(r_int_nrofil, 14) = Format(grd_Listad.TextMatrix(r_int_nroaux, 18), "###,###,##0.00")
         
         r_int_nrofil = r_int_nrofil + 1
      Next
      .Range(.Cells(r_int_nrofil, 12), .Cells(r_int_nrofil, 14)).FormulaR1C1 = "=SUM(R[-" & r_int_nrofil - 4 & "]C:R[-1]C)"
      .Range(.Cells(r_int_nrofil, 12), .Cells(r_int_nrofil, 14)).Font.Bold = True
      .Cells(r_int_nrofil, 12).NumberFormat = "#,##0.00"
      .Cells(r_int_nrofil, 13).NumberFormat = "0"
      .Cells(r_int_nrofil, 14).NumberFormat = "#,##0.00"
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_Situac_Click()
   Screen.MousePointer = 11
   Call fs_Carga_grid
   Screen.MousePointer = 0
End Sub

Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_PriVct_Click()
   If Len(Trim(pnl_Tit_PriVct.Tag)) = 0 Or pnl_Tit_PriVct.Tag = "D" Then
      pnl_Tit_PriVct.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 13, "C")
   Else
      pnl_Tit_PriVct.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 13, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumMes_Click()
   If Len(Trim(pnl_Tit_NumMes.Tag)) = 0 Or pnl_Tit_NumMes.Tag = "D" Then
      pnl_Tit_NumMes.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_NumMes.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Tit_MtoAho_Click()
   If Len(Trim(pnl_Tit_MtoAho.Tag)) = 0 Or pnl_Tit_MtoAho.Tag = "D" Then
      pnl_Tit_MtoAho.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Tit_MtoAho.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

Private Sub pnl_Tit_Situac_Click()
   If Len(Trim(pnl_Tit_Situac.Tag)) = 0 Or pnl_Tit_Situac.Tag = "D" Then
      pnl_Tit_Situac.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 7, "C")
   Else
      pnl_Tit_Situac.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 7, "C-")
   End If
End Sub

Private Sub pnl_Tit_Consejero_Click()
   If Len(Trim(pnl_Tit_Consejero.Tag)) = 0 Or pnl_Tit_Consejero.Tag = "D" Then
      pnl_Tit_Consejero.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "C")
   Else
      pnl_Tit_Consejero.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "C-")
   End If
End Sub
