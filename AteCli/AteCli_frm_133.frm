VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_PryNVi_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   2460
   ClientTop       =   3315
   ClientWidth     =   11625
   Icon            =   "AteCli_frm_133.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4875
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
      _ExtentY        =   8599
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
         Height          =   2085
         Left            =   30
         TabIndex        =   18
         Top             =   2730
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   3678
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
         Begin VB.TextBox txt_Telef1 
            Height          =   315
            Left            =   1860
            MaxLength       =   12
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   1710
            Width           =   1640
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   1860
            TabIndex        =   10
            Text            =   "cmb_DstDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   1860
            TabIndex        =   8
            Text            =   "cmb_DptDir"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   1860
            MaxLength       =   120
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_NumFax 
            Height          =   315
            Left            =   8130
            MaxLength       =   12
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   1710
            Width           =   1640
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8130
            MaxLength       =   250
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8130
            TabIndex        =   9
            Text            =   "cmb_PrvDir"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8130
            MaxLength       =   120
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_IntDpt 
            Height          =   315
            Left            =   9780
            MaxLength       =   15
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.TextBox txt_NumVia 
            Height          =   315
            Left            =   8130
            MaxLength       =   15
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_Telef2 
            Height          =   315
            Left            =   3510
            MaxLength       =   12
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   1710
            Width           =   1670
         End
         Begin VB.Label Label27 
            Caption         =   "Teléfono:"
            Height          =   285
            Left            =   60
            TabIndex        =   29
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label Label26 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   60
            TabIndex        =   28
            Top             =   1380
            Width           =   1605
         End
         Begin VB.Label Label24 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   60
            TabIndex        =   27
            Top             =   1050
            Width           =   1305
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label Label20 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   60
            TabIndex        =   25
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label29 
            Caption         =   "Fax:"
            Height          =   315
            Left            =   6120
            TabIndex        =   24
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label Label28 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   6120
            TabIndex        =   23
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label25 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   6120
            TabIndex        =   22
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label23 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   6120
            TabIndex        =   21
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label21 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   6120
            TabIndex        =   20
            Top             =   390
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Vía:"
            Height          =   285
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   30
         Top             =   1920
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.TextBox txt_RazSoc 
            Height          =   315
            Left            =   1860
            MaxLength       =   250
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   60
            Width           =   9615
         End
         Begin VB.TextBox txt_NomCom 
            Height          =   315
            Left            =   1860
            MaxLength       =   250
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   9615
         End
         Begin VB.Label Label32 
            Caption         =   "Razón Social:"
            Height          =   285
            Left            =   60
            TabIndex        =   32
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Left            =   60
            TabIndex        =   31
            Top             =   390
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   33
         Top             =   30
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
            TabIndex        =   34
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Gestión de Proyectos No Vinculados"
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
            Picture         =   "AteCli_frm_133.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel15 
         Height          =   645
         Left            =   30
         TabIndex        =   35
         Top             =   750
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_133.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10950
            Picture         =   "AteCli_frm_133.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Left            =   30
         TabIndex        =   36
         Top             =   1440
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   767
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
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   315
            Left            =   1860
            TabIndex        =   37
            Top             =   60
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1 - 07522154"
            ForeColor       =   32768
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
         Begin VB.Label Label1 
            Caption         =   "Docum. de Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   38
            Top             =   60
            Width           =   1725
         End
      End
   End
End
Attribute VB_Name = "frm_PryNVi_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_FlgCmb     As Integer
Dim l_str_DptDir     As String
Dim l_str_PrvDir     As String
Dim l_str_DstDir     As String

Private Sub cmb_DptDir_Change()
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_Click()
   If cmb_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvDir.Clear
         cmb_DstDir.Clear
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_DptDir_GotFocus()
   Call SendMessage(cmb_DptDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   l_int_FlgCmb = True
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptDir, l_str_DptDir)
      l_int_FlgCmb = True
      
      cmb_PrvDir.Clear
      cmb_DstDir.Clear
      If cmb_DptDir.ListIndex > -1 Then
         l_str_DptDir = ""
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvDir)
   End If
End Sub

Private Sub cmb_DptDir_LostFocus()
   Call SendMessage(cmb_DptDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_PrvDir_Change()
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_Click()
   If cmb_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstDir.Clear
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstDir)
      End If
   End If
End Sub

Private Sub cmb_PrvDir_GotFocus()
   Call SendMessage(cmb_PrvDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   l_int_FlgCmb = True
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvDir, l_str_PrvDir)
      l_int_FlgCmb = True
      
      cmb_DstDir.Clear
      If cmb_PrvDir.ListIndex > -1 Then
         l_str_DstDir = ""
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstDir)
   End If
End Sub

Private Sub cmb_PrvDir_LostFocus()
   Call SendMessage(cmb_PrvDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_DstDir_Change()
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_Click()
   If cmb_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Refere)
      End If
   End If
End Sub

Private Sub cmb_DstDir_GotFocus()
   Call SendMessage(cmb_DstDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   l_int_FlgCmb = True
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstDir, l_str_DstDir)
      l_int_FlgCmb = True
      
      If cmb_DstDir.ListIndex > -1 Then
         l_str_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Refere)
   End If
End Sub

Private Sub cmb_DstDir_LostFocus()
   Call SendMessage(cmb_DstDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_TipVia_Click()
   Call gs_SetFocus(txt_NomVia)
End Sub

Private Sub cmb_TipVia_KeyPress(KeyAscii As Integer)
   Call cmb_TipVia_Click
End Sub

Private Sub cmb_TipZon_Click()
   Call gs_SetFocus(txt_NomZon)
End Sub

Private Sub cmb_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_TipZon_Click
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_RazSoc.Text)) = 0 Then
      MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RazSoc)
      Exit Sub
   End If
   If Len(Trim(txt_NomCom.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomCom)
      Exit Sub
   End If
   If cmb_TipVia.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipVia)
      Exit Sub
   End If
   If Len(Trim(txt_NomVia.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre de la Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomVia)
      Exit Sub
   End If
   If Len(Trim(txt_NumVia.Text)) = 0 Then
      MsgBox "Debe ingresar el Número en la Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumVia)
      Exit Sub
   End If
   If cmb_TipZon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipZon)
      Exit Sub
   End If
   If cmb_TipZon.ItemData(cmb_TipZon.ListIndex) <> 12 Then
      If Len(Trim(txt_NomVia.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de la Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomZon)
         Exit Sub
      End If
   End If
   If cmb_DptDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Departamento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DptDir)
      Exit Sub
   End If
   If cmb_PrvDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Provincia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PrvDir)
      Exit Sub
   End If
   If cmb_DstDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Distrito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DstDir)
      Exit Sub
   End If
   If Len(Trim(txt_Telef1.Text)) = 0 Then
      MsgBox "Debe ingresar el Teléfono.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Telef1)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando Información del Cliente
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_EMP_DATGEN_PRYNVI ("
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & txt_RazSoc.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NomCom.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_IntDpt.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Telef1.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Telef2.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumFax.Text & "', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ")"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop

   MsgBox "Los datos fueron registrados correctamente.", vbInformation, modgen_g_str_NomPlt
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   pnl_DocIde.Caption = moddat_gf_Consulta_ParDes("232", moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc
   
   Call fs_Inicio
   Call fs_Limpia
   
   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_str_NumDoc & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         txt_RazSoc.Text = Trim(g_rst_Princi!DATGEN_RAZSOC & "")
         txt_NomCom.Text = Trim(g_rst_Princi!DATGEN_NOMCOM & "")
         
         Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!DatGen_TipVia)
         txt_NomVia.Text = Trim(g_rst_Princi!DatGen_NomVia & "")
         txt_NumVia.Text = Trim(g_rst_Princi!DatGen_numVia & "")
         txt_IntDpt.Text = Trim(g_rst_Princi!DATGEN_INTDPT & "")
         Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!DatGen_TipZon)
         txt_NomZon.Text = Trim(g_rst_Princi!DatGen_NomZon & "")
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(Format(Trim(g_rst_Princi!DatGen_Ubigeo & ""), "000000"), 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(Format(Trim(g_rst_Princi!DatGen_Ubigeo & ""), "000000"), 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(Format(Trim(g_rst_Princi!DatGen_Ubigeo & ""), "000000"), 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir, Left(Format(Trim(g_rst_Princi!DatGen_Ubigeo & ""), "000000"), 2), Mid(Format(Trim(g_rst_Princi!DatGen_Ubigeo & ""), "000000"), 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(Format(Trim(g_rst_Princi!DatGen_Ubigeo & ""), "000000"), 2)))
         txt_Refere.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
         txt_Telef1.Text = Trim(g_rst_Princi!DATGEN_TELEF1 & "")
         txt_Telef2.Text = Trim(g_rst_Princi!DATGEN_TELEF2 & "")
         txt_NumFax.Text = Trim(g_rst_Princi!DATGEN_NUMFAX & "")
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Call gs_CentraForm(Me)
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_DptDir)
End Sub

Private Sub fs_Limpia()
   txt_RazSoc.Text = ""
   txt_NomCom.Text = ""
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_NumVia.Text = ""
   txt_IntDpt.Text = ""
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Refere.Text = ""
   txt_Telef1.Text = ""
   txt_Telef2.Text = ""
   txt_NumFax.Text = ""
End Sub

Private Sub txt_NomCom_GotFocus()
   Call gs_SelecTodo(txt_NomCom)
End Sub

Private Sub txt_NomCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,@#$%&;:()/º")
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomZon_GotFocus()
   Call gs_SelecTodo(txt_NomZon)
End Sub

Private Sub txt_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NumVia_GotFocus()
   Call gs_SelecTodo(txt_NumVia)
End Sub

Private Sub txt_NumVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_IntDpt_GotFocus()
   Call gs_SelecTodo(txt_IntDpt)
End Sub

Private Sub txt_IntDpt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_RazSoc)
End Sub

Private Sub txt_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,@#$%&;:()/º")
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Telef1_GotFocus()
   Call gs_SelecTodo(txt_Telef1)
End Sub

Private Sub txt_Telef1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Telef2_GotFocus()
   Call gs_SelecTodo(txt_Telef2)
End Sub

Private Sub txt_Telef2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumFax)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumFax_GotFocus()
   Call gs_SelecTodo(txt_NumFax)
End Sub

Private Sub txt_NumFax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

