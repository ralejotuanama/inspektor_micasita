VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_PryAsig_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13935
   Icon            =   "AteCli_frm_189.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5715
      Left            =   30
      TabIndex        =   9
      Top             =   1350
      Width           =   13875
      _Version        =   65536
      _ExtentX        =   24474
      _ExtentY        =   10081
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   5565
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   13815
         _Version        =   65536
         _ExtentX        =   24368
         _ExtentY        =   9816
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   585
            Left            =   120
            TabIndex        =   17
            Top             =   4890
            Width           =   13575
            _Version        =   65536
            _ExtentX        =   23945
            _ExtentY        =   1032
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
            Begin VB.TextBox txt_NomProy 
               Height          =   315
               Left            =   1710
               TabIndex        =   2
               Text            =   "Text1"
               Top             =   150
               Width           =   4665
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Buscar por Proyecto:"
               Height          =   195
               Left            =   150
               TabIndex        =   18
               Top             =   180
               Width           =   1485
            End
         End
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            Left            =   5340
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   4365
         End
         Begin MSFlexGridLib.MSFlexGrid grdListaProy 
            Height          =   3885
            Left            =   120
            TabIndex        =   1
            Top             =   960
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   6853
            _Version        =   393216
            FixedCols       =   0
            BackColorSel    =   32768
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdEliminar 
            Height          =   585
            Left            =   6600
            Picture         =   "AteCli_frm_189.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Eliminar Proyecto"
            Top             =   2850
            Width           =   585
         End
         Begin VB.CommandButton cmdAgregar 
            Height          =   585
            Left            =   6600
            Picture         =   "AteCli_frm_189.frx":08D6
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Asignar Proyecto"
            Top             =   2160
            Width           =   585
         End
         Begin MSFlexGridLib.MSFlexGrid grdListaAsig 
            Height          =   3885
            Left            =   7320
            TabIndex        =   5
            Top             =   960
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   6853
            _Version        =   393216
            FixedCols       =   0
            BackColorSel    =   32768
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Consejero Hipotecario:"
            Height          =   195
            Left            =   3660
            TabIndex        =   16
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Lista de Proyectos Asignados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7350
            TabIndex        =   15
            Top             =   720
            Width           =   6255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Lista de Proyectos Disponibles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   150
            TabIndex        =   14
            Top             =   720
            Width           =   6255
         End
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   675
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   13905
      _Version        =   65536
      _ExtentX        =   24527
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
         TabIndex        =   12
         Top             =   60
         Width           =   5445
         _Version        =   65536
         _ExtentX        =   9604
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "Asignación de Proyectos a Consejeros Hipotecarios"
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
         Picture         =   "AteCli_frm_189.frx":11A0
         Top             =   60
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   690
      Width           =   13905
      _Version        =   65536
      _ExtentX        =   24527
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
      Begin VB.CommandButton cmd_Limpiar 
         Height          =   585
         Left            =   630
         Picture         =   "AteCli_frm_189.frx":14AA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpiar Datos"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Buscar 
         Height          =   585
         Left            =   60
         Picture         =   "AteCli_frm_189.frx":17B4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar Datos"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salir 
         Height          =   585
         Left            =   13260
         Picture         =   "AteCli_frm_189.frx":1ABE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
   End
End
Attribute VB_Name = "frm_PryAsig_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_ConHip()   As moddat_tpo_Genera

Private Sub Busqueda_Proyecto(NombreProyecto As String)
   Dim r_int_Cont As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DATGEN_CODIGO, (TRIM(DATGEN_TITULO) || ' - ' || TRIM(MNT_PARDES.PARDES_DESCRI)) DATGEN_TITULO "
   g_str_Parame = g_str_Parame & "  FROM PRY_DATGEN "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES ON MNT_PARDES.PARDES_CODGRP = 513 AND PRY_DATGEN.DATGEN_CODBCO = MNT_PARDES.PARDES_CODITE "
   g_str_Parame = g_str_Parame & "WHERE DATGEN_PRYMCS = 2 "
   g_str_Parame = g_str_Parame & "  AND (TRIM(DATGEN_TITULO) || ' - ' || TRIM(MNT_PARDES.PARDES_DESCRI)) LIKE '%" & NombreProyecto & "%'"
   g_str_Parame = g_str_Parame & "ORDER BY 2"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_Cont = 1
   grdListaProy.Rows = 1
      
   If g_rst_Princi.EOF And g_rst_Princi.BOF Then
      cmdAgregar.Enabled = False
   Else
      cmdAgregar.Enabled = True
   End If
   
   Do While Not g_rst_Princi.EOF
      grdListaProy.Rows = grdListaProy.Rows + 1
      grdListaProy.SelectionMode = flexSelectionByRow
      grdListaProy.TextMatrix(r_int_Cont, 0) = g_rst_Princi!DATGEN_CODIGO
      grdListaProy.TextMatrix(r_int_Cont, 1) = g_rst_Princi!DATGEN_TITULO
      
      r_int_Cont = r_int_Cont + 1
      g_rst_Princi.MoveNext
   Loop
End Sub

Private Sub fs_Inicio()
   With grdListaProy
      .ColWidth(0) = 0
      .ColWidth(1) = 6050
      .FixedAlignment(1) = flexAlignCenterCenter
      
      .TextMatrix(0, 0) = "Codigo"
      .TextMatrix(0, 1) = "Nombre de Proyectos"
   End With

   With grdListaAsig
      .ColWidth(0) = 0
      .ColWidth(1) = 6050
      .FixedAlignment(1) = flexAlignCenterCenter
      
      .TextMatrix(0, 0) = "Codigo"
      .TextMatrix(0, 1) = "Proyectos Asignado a Consejero"
   End With
   
   cmdAgregar.Enabled = False
   cmdEliminar.Enabled = False
   txt_NomProy.Enabled = False
   txt_NomProy.Text = ""
End Sub

Private Sub cmd_Buscar_Click()
   Dim r_int_Cont As Integer
   
   If cmb_ConHip.Text = "" Then
      MsgBox "Seleccione Consejero Hipotecario para iniciar Asignacion.", vbExclamation, modgen_g_str_NomPlt
      cmb_ConHip.SetFocus
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DATGEN_CODIGO, (TRIM(DATGEN_TITULO) || ' - ' || TRIM(MNT_PARDES.PARDES_DESCRI)) DATGEN_TITULO "
   g_str_Parame = g_str_Parame & "FROM PRY_DATGEN LEFT JOIN MNT_PARDES ON MNT_PARDES.PARDES_CODGRP = 513 AND PRY_DATGEN.DATGEN_CODBCO=MNT_PARDES.PARDES_CODITE "
   g_str_Parame = g_str_Parame & "WHERE DATGEN_PRYMCS = 2 "
   g_str_Parame = g_str_Parame & "ORDER BY 2"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_Cont = 1
   grdListaProy.Rows = 1
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grdListaProy.Rows = grdListaProy.Rows + 1
      grdListaProy.SelectionMode = flexSelectionByRow
      grdListaProy.TextMatrix(r_int_Cont, 0) = g_rst_Princi!DATGEN_CODIGO
      grdListaProy.TextMatrix(r_int_Cont, 1) = g_rst_Princi!DATGEN_TITULO
      
      r_int_Cont = r_int_Cont + 1
      g_rst_Princi.MoveNext
   Loop

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DATGEN_CODIGO,(TRIM(DATGEN_TITULO) || ' - ' || TRIM(MNT_PARDES.PARDES_DESCRI)) DATGEN_TITULO,ASGCON_CONHIP,DATGEN_PRYMCS "
   g_str_Parame = g_str_Parame & " FROM PRY_DATGEN LEFT JOIN MNT_PARDES ON MNT_PARDES.PARDES_CODGRP = 513 AND PRY_DATGEN.DATGEN_CODBCO=MNT_PARDES.PARDES_CODITE "
   g_str_Parame = g_str_Parame & " LEFT JOIN PRY_ASGCON ON PRY_DATGEN.DATGEN_CODIGO=PRY_ASGCON.ASGCON_CODPRY"
   g_str_Parame = g_str_Parame & " WHERE DATGEN_PRYMCS = '2' AND ASGCON_CONHIP = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   r_int_Cont = 1
   grdListaAsig.Rows = 1

   If g_rst_Princi.EOF And g_rst_Princi.BOF Then grdListaAsig.Rows = 2
   Do While Not g_rst_Princi.EOF
      grdListaAsig.Rows = grdListaAsig.Rows + 1
      grdListaAsig.TextMatrix(r_int_Cont, 0) = g_rst_Princi!DATGEN_CODIGO
      grdListaAsig.TextMatrix(r_int_Cont, 1) = g_rst_Princi!DATGEN_TITULO

      r_int_Cont = r_int_Cont + 1
      g_rst_Princi.MoveNext

   Loop
   
   cmdAgregar.Enabled = True
   cmdEliminar.Enabled = True
   cmb_ConHip.Enabled = False
   txt_NomProy.Enabled = True
End Sub

Private Sub cmd_Limpiar_Click()
   grdListaProy.Rows = 2
   grdListaAsig.Rows = 2
   grdListaProy.TextMatrix(1, 0) = ""
   grdListaProy.TextMatrix(1, 1) = ""
   grdListaAsig.TextMatrix(1, 0) = ""
   grdListaAsig.TextMatrix(1, 1) = ""
   
   cmb_ConHip.ListIndex = -1
   cmdAgregar.Enabled = False
   cmdEliminar.Enabled = False
   cmb_ConHip.Enabled = True
   txt_NomProy.Enabled = False
   txt_NomProy.Text = ""
   cmb_ConHip.SetFocus
End Sub

Private Sub cmd_Salir_Click()
   Unload Me
End Sub

Private Sub cmdAgregar_Click()
   Dim r_rst_Grabar As ADODB.Recordset
   Dim g_str_Parame As String
   
   moddat_g_int_CntErr = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame + "SELECT * FROM PRY_ASGCON "
   g_str_Parame = g_str_Parame + "WHERE ASGCON_CODPRY='" & Trim(grdListaProy.TextMatrix(grdListaProy.Row, 0)) & "'"
   g_str_Parame = g_str_Parame + " AND ASGCON_CONHIP='" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 3) Then
      Exit Sub
   End If
      
   If Not (r_rst_Grabar.EOF And r_rst_Grabar.BOF) Then
      MsgBox "El Proyecto seleccionado ya se encuentra registrado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   g_str_Parame = ""
   g_str_Parame = g_str_Parame + "INSERT INTO PRY_ASGCON (ASGCON_CODPRY,ASGCON_CONHIP) "
   g_str_Parame = g_str_Parame + " VALUES ('" & Trim(grdListaProy.TextMatrix(grdListaProy.Row, 0)) & "', "
   g_str_Parame = g_str_Parame + " '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "')"
                
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
      moddat_g_int_CntErr = moddat_g_int_CntErr + 1
   Else
      moddat_g_int_FlgGOK = True
   End If
   
   cmd_Buscar_Click
End Sub

Private Sub cmdEliminar_Click()
   Dim r_rst_Grabar As ADODB.Recordset
   Dim g_str_Parame As String
   
   moddat_g_int_CntErr = 0
   
   If Trim(grdListaAsig.TextMatrix(grdListaAsig.Row, 0)) = "" Then Exit Sub
   
   If MsgBox("¿Esta seguro de Eliminar el registro seleccionado?.", vbYesNo + vbQuestion + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + " DELETE FROM PRY_ASGCON "
      g_str_Parame = g_str_Parame + " WHERE ASGCON_CODPRY='" & Trim(grdListaAsig.TextMatrix(grdListaAsig.Row, 0)) & "' "
      g_str_Parame = g_str_Parame + " AND ASGCON_CONHIP='" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "'"
      
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
   
      cmd_Buscar_Click
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicio
   Call gs_CentraForm(Me)
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
   
   Screen.MousePointer = 0
End Sub

Private Sub grdListaAsig_SelChange()
   If grdListaAsig.Rows > 2 Then
      grdListaAsig.RowSel = grdListaAsig.Row
   End If
End Sub

Private Sub grdListaProy_SelChange()
   If grdListaProy.Rows > 2 Then
      grdListaProy.RowSel = grdListaProy.Row
   End If
End Sub

Private Sub txt_NomProy_GotFocus()
   Call gs_SelecTodo(txt_NomProy)
End Sub

Private Sub txt_NomProy_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call Busqueda_Proyecto(txt_NomProy.Text)
      Call gs_SetFocus(grdListaProy)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomProy_LostFocus()
   If Len(Trim(txt_NomProy.Text)) > 0 Then
      Call Busqueda_Proyecto(txt_NomProy.Text)
      Call gs_SetFocus(grdListaProy)
   End If
End Sub
