VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_MntCli_67 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8115
   Icon            =   "AteCli_frm_578.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      _Version        =   65536
      _ExtentX        =   14314
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   2235
         Left            =   30
         TabIndex        =   1
         Top             =   3420
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
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
         BevelOuter      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2115
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   7905
            _ExtentX        =   13944
            _ExtentY        =   3731
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   750
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7410
            Picture         =   "AteCli_frm_578.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_578.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_578.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
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
            TabIndex        =   8
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Actualización de Datos de Clientes"
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
            Picture         =   "AteCli_frm_578.frx":0A62
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1095
         Left            =   30
         TabIndex        =   9
         Top             =   2280
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   1931
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
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   585
            Left            =   6750
            Picture         =   "AteCli_frm_578.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Buscar Clientes"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_LimBus 
            Height          =   585
            Left            =   7380
            Picture         =   "AteCli_frm_578.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1725
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   390
            Width           =   1725
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   30
            TabIndex        =   15
            Top             =   720
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   18
         Top             =   1470
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   1349
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
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   60
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   22
            Top             =   390
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   1845
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_67"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
      
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
      txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
   End If

   moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   moddat_g_str_TipDoc = cmb_TipDoc.Text
   moddat_g_str_NumDoc = txt_NumDoc.Text

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      moddat_g_int_FlgGrb = 1
   Else
      moddat_g_int_FlgGrb = 2
   End If
   
   frm_MntCli_68.Show 1
End Sub

Private Sub cmd_BusCli_Click()
   Dim r_str_ApePat  As String
   Dim r_str_ApeMat  As String
   Dim r_str_Nombre  As String

   If Len(Trim(txt_ApePat)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_Listad)
   r_str_ApePat = txt_ApePat.Text & "%"
   r_str_ApeMat = txt_ApeMat.Text & "%"
   r_str_Nombre = txt_Nombre.Text & "%"
   
   g_str_Parame = "SELECT * FROM CLI_BUSCLI WHERE "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_APEPAT) LIKE '" & r_str_ApePat & "' AND "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_APEMAT) LIKE '" & r_str_ApeMat & "' AND "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_NOMBRE) LIKE '" & r_str_Nombre & "' ORDER BY "
   g_str_Parame = g_str_Parame & "BUSCLI_APEPAT ASC, "
   g_str_Parame = g_str_Parame & "BUSCLI_APEMAT ASC, "
   g_str_Parame = g_str_Parame & "BUSCLI_NOMBRE ASC "
   
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
      grd_Listad.Text = CStr(g_rst_Princi!BUSCLI_TIPDOC) & "-" & Trim(g_rst_Princi!BUSCLI_NUMDOC & "")
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!BUSCLI_APEPAT & "") & " " & Trim(g_rst_Princi!BUSCLI_APEMAT & "") & " " & Trim(g_rst_Princi!BUSCLI_NOMBRE & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   Call gs_RefrescaGrid(grd_Listad)
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_LimBus_Click()
   Call fs_Limpia_BusAlf
   Call gs_SetFocus(txt_ApePat)
End Sub

Private Sub cmd_Limpia_Click()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicio
   Call cmd_Limpia_Click
   Call fs_Limpia_BusAlf
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   grd_Listad.ColWidth(0) = 2000
   grd_Listad.ColWidth(1) = 6000
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
End Sub

Private Sub fs_Limpia_BusAlf()
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub grd_Listad_DblClick()
   Dim r_int_TipDoc     As Integer
   Dim r_str_NumDoc     As String

   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 0
      
      r_int_TipDoc = CInt(Left(grd_Listad.Text, 1))
      r_str_NumDoc = Mid(grd_Listad.Text, 3)
   
      Call gs_RefrescaGrid(grd_Listad)
      
      Call gs_BuscarCombo_Item(cmb_TipDoc, r_int_TipDoc)
      txt_NumDoc.Text = r_str_NumDoc
      
      Call cmd_Buscar_Click
      Call gs_SetFocus(grd_Listad)
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusCli)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 4:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub
