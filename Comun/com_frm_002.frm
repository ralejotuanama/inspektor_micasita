VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_IdeUsu_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2895
   ClientLeft      =   8535
   ClientTop       =   3240
   ClientWidth     =   4785
   Icon            =   "com_frm_002.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2895
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4785
      _Version        =   65536
      _ExtentX        =   8440
      _ExtentY        =   5106
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
         Height          =   915
         Left            =   30
         TabIndex        =   11
         Top             =   1920
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   1614
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
         Begin VB.TextBox txt_Nuevo2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1950
            MaxLength       =   30
            PasswordChar    =   "#"
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   420
            Width           =   2685
         End
         Begin VB.TextBox txt_Nuevo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1950
            MaxLength       =   30
            PasswordChar    =   "#"
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   60
            Width           =   2685
         End
         Begin VB.Label Label3 
            Caption         =   "Contraseña Nueva (Confirmación):"
            Height          =   465
            Left            =   60
            TabIndex        =   13
            Top             =   420
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Contraseña Nueva:"
            Height          =   345
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1545
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   435
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
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
         Begin VB.TextBox txt_Actual 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1950
            MaxLength       =   30
            PasswordChar    =   "#"
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   60
            Width           =   2685
         End
         Begin VB.Label Label1 
            Caption         =   "Contraseña Actual:"
            Height          =   345
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   1545
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   450
            Left            =   690
            TabIndex        =   9
            Top             =   60
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   794
            _StockProps     =   15
            Caption         =   "Cambio de Contraseña"
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
            Picture         =   "com_frm_002.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
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
            Picture         =   "com_frm_002.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4050
            Picture         =   "com_frm_002.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_IdeUsu_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Grabar_Click()
'   If Len(Trim(txt_Actual.Text)) = 0 Then
'      MsgBox "La contraseña actual esta vacía.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_Actual)
'      Exit Sub
'   End If
'   If Len(Trim(txt_Nuevo1.Text)) = 0 Then
'      MsgBox "La contraseña nueva ingresada no es válida.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_Nuevo1)
'      Exit Sub
'   End If
'   If Len(Trim(txt_Nuevo2.Text)) = 0 Then
'      MsgBox "La confirmación de la contraseña nueva ingresada no es válida.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_Nuevo2)
'      Exit Sub
'   End If
'   If txt_Actual.Text = txt_Nuevo1.Text Then
'      MsgBox "La contraseña nueva no puede ser igual a la actual.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_Nuevo1)
'      Exit Sub
'   End If
'   If txt_Nuevo1.Text <> txt_Nuevo2.Text Then
'      MsgBox "La confirmación de la contraseña no coincide con la contraseña nueva.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_Nuevo2)
'      Exit Sub
'   End If
'
'   g_str_Parame = "SELECT USUMAE_CONTRA FROM SEG_USUMAE WHERE USUMAE_CODIGO = '" & modgen_g_str_CodUsu & "' AND USUMAE_SITUAC = 1"
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'      Exit Sub
'   End If
'
'   g_rst_Princi.MoveFirst
'
'   If gf_Seg_Desenc(Trim(g_rst_Princi!USUMAE_CONTRA)) <> txt_Actual.Text Then
'      g_rst_Princi.Close
'      Set g_rst_Princi = Nothing
'
'      MsgBox "La Contraseña Actual es incorrecta.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_Actual)
'      Exit Sub
'   End If
'
'   g_rst_Princi.Close
'   Set g_rst_Princi = Nothing
'
'   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
'      Exit Sub
'   End If
'
'   moddat_g_int_FlgGOK = False
'   moddat_g_int_CntErr = 0
'
'   Do While moddat_g_int_FlgGOK = False
'      g_str_Parame = "USP_SEG_USUMAE_CAMBIOCLAVE ("
'      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
'      g_str_Parame = g_str_Parame & "'" & gf_Seg_Encrip(txt_Nuevo1.Text) & "', "
'
'      'Datos de Auditoria
'      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
'      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
'      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
'      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
'
'      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
'         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
'      Else
'         moddat_g_int_FlgGOK = True
'      End If
'
'      If moddat_g_int_CntErr = 6 Then
'         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
'            Exit Sub
'         Else
'            moddat_g_int_CntErr = 0
'         End If
'      End If
'   Loop
'
'   MsgBox "El cambio de clave se realizó con éxito.", vbInformation, modgen_g_str_NomPlt
'   Unload Me
   
   If gf_Seg_CamClave(txt_Actual, txt_Nuevo1, txt_Nuevo2) = True Then
      Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   txt_Actual.Text = ""
   txt_Nuevo1.Text = ""
   txt_Nuevo2.Text = ""
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub txt_Actual_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nuevo1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS)
   End If
End Sub

Private Sub txt_Nuevo1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nuevo2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS)
   End If
End Sub

Private Sub txt_Nuevo2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS)
   End If
End Sub


