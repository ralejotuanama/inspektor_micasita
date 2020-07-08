VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_IngSol_06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   4260
   ClientLeft      =   1785
   ClientTop       =   2850
   ClientWidth     =   11640
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4245
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   7488
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   3465
         Left            =   30
         TabIndex        =   16
         Top             =   720
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   6112
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
         Begin VB.CommandButton cmd_Acepta 
            Height          =   675
            Left            =   10020
            Picture         =   "AteCli_frm_005.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Aceptar Datos"
            Top             =   2700
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10740
            Picture         =   "AteCli_frm_005.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Salir de la Opción"
            Top             =   2700
            Width           =   675
         End
         Begin Threed.SSPanel SSPanel24 
            Height          =   90
            Left            =   30
            TabIndex        =   17
            Top             =   2580
            Width           =   11415
            _Version        =   65536
            _ExtentX        =   20135
            _ExtentY        =   159
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
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin TabDlg.SSTab tab_Princi 
            Height          =   2475
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   4366
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Referencia Familiar"
            TabPicture(0)   =   "AteCli_frm_005.frx":074C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "txt_Fam_ApePat"
            Tab(0).Control(1)=   "cmb_Fam_TipPar"
            Tab(0).Control(2)=   "txt_Fam_ApeMat"
            Tab(0).Control(3)=   "txt_Fam_Nombre"
            Tab(0).Control(4)=   "txt_Fam_Telefo"
            Tab(0).Control(5)=   "txt_Fam_Celula"
            Tab(0).Control(6)=   "Label19"
            Tab(0).Control(7)=   "Label2"
            Tab(0).Control(8)=   "Label3"
            Tab(0).Control(9)=   "Label14"
            Tab(0).Control(10)=   "Label17"
            Tab(0).Control(11)=   "Label22"
            Tab(0).ControlCount=   12
            TabCaption(1)   =   "Referencia No Familiar"
            TabPicture(1)   =   "AteCli_frm_005.frx":0768
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Label4"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label5"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "Label7"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "Label8"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "Label9"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "Label10"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "cmd_ActEco"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "txt_NFa_Celula"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "txt_NFa_Telefo"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "txt_NFa_Nombre"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "txt_NFa_ApeMat"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "cmb_NFa_TipPar"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).Control(12)=   "txt_NFa_ApePat"
            Tab(1).Control(12).Enabled=   0   'False
            Tab(1).ControlCount=   13
            Begin VB.TextBox txt_NFa_ApePat 
               Height          =   315
               Left            =   2010
               MaxLength       =   30
               TabIndex        =   7
               Text            =   "Text1"
               Top             =   750
               Width           =   3315
            End
            Begin VB.ComboBox cmb_NFa_TipPar 
               Height          =   315
               Left            =   2010
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   420
               Width           =   3315
            End
            Begin VB.TextBox txt_NFa_ApeMat 
               Height          =   315
               Left            =   2010
               MaxLength       =   30
               TabIndex        =   8
               Text            =   "Text1"
               Top             =   1080
               Width           =   3315
            End
            Begin VB.TextBox txt_NFa_Nombre 
               Height          =   315
               Left            =   2010
               MaxLength       =   30
               TabIndex        =   9
               Text            =   "Text1"
               Top             =   1410
               Width           =   3315
            End
            Begin VB.TextBox txt_NFa_Telefo 
               Height          =   315
               Left            =   2010
               MaxLength       =   12
               TabIndex        =   10
               Text            =   "Text1"
               Top             =   1740
               Width           =   3315
            End
            Begin VB.TextBox txt_NFa_Celula 
               Height          =   315
               Left            =   2010
               MaxLength       =   12
               TabIndex        =   11
               Text            =   "Text1"
               Top             =   2070
               Width           =   3315
            End
            Begin VB.CommandButton cmd_ActEco 
               Height          =   675
               Left            =   15990
               Picture         =   "AteCli_frm_005.frx":0784
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "Actividades Económicas"
               Top             =   7140
               Width           =   675
            End
            Begin VB.TextBox txt_Fam_ApePat 
               Height          =   315
               Left            =   -72990
               MaxLength       =   30
               TabIndex        =   1
               Text            =   "Text1"
               Top             =   750
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Fam_TipPar 
               Height          =   315
               Left            =   -72990
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   420
               Width           =   3315
            End
            Begin VB.TextBox txt_Fam_ApeMat 
               Height          =   315
               Left            =   -72990
               MaxLength       =   30
               TabIndex        =   2
               Text            =   "Text1"
               Top             =   1080
               Width           =   3315
            End
            Begin VB.TextBox txt_Fam_Nombre 
               Height          =   315
               Left            =   -72990
               MaxLength       =   30
               TabIndex        =   3
               Text            =   "Text1"
               Top             =   1410
               Width           =   3315
            End
            Begin VB.TextBox txt_Fam_Telefo 
               Height          =   315
               Left            =   -72990
               MaxLength       =   12
               TabIndex        =   4
               Text            =   "Text1"
               Top             =   1740
               Width           =   3315
            End
            Begin VB.TextBox txt_Fam_Celula 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -72990
               MaxLength       =   12
               TabIndex        =   5
               Text            =   "Text1"
               Top             =   2070
               Width           =   3315
            End
            Begin VB.Label Label10 
               Caption         =   "Tipo Parentesco:"
               Height          =   315
               Left            =   90
               TabIndex        =   31
               Top             =   420
               Width           =   1905
            End
            Begin VB.Label Label9 
               Caption         =   "Apellido Paterno:"
               Height          =   315
               Left            =   90
               TabIndex        =   30
               Top             =   750
               Width           =   1905
            End
            Begin VB.Label Label8 
               Caption         =   "Apellido Materno:"
               Height          =   315
               Left            =   90
               TabIndex        =   29
               Top             =   1080
               Width           =   1905
            End
            Begin VB.Label Label7 
               Caption         =   "Nombres:"
               Height          =   315
               Left            =   90
               TabIndex        =   28
               Top             =   1410
               Width           =   1905
            End
            Begin VB.Label Label5 
               Caption         =   "Teléfono:"
               Height          =   315
               Left            =   90
               TabIndex        =   27
               Top             =   1740
               Width           =   1905
            End
            Begin VB.Label Label4 
               Caption         =   "Celular:"
               Height          =   315
               Left            =   90
               TabIndex        =   26
               Top             =   2070
               Width           =   1905
            End
            Begin VB.Label Label19 
               Caption         =   "Tipo Parentesco:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   25
               Top             =   420
               Width           =   1905
            End
            Begin VB.Label Label2 
               Caption         =   "Apellido Paterno:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   24
               Top             =   750
               Width           =   1905
            End
            Begin VB.Label Label3 
               Caption         =   "Apellido Materno:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   23
               Top             =   1080
               Width           =   1905
            End
            Begin VB.Label Label14 
               Caption         =   "Nombres:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   22
               Top             =   1410
               Width           =   1905
            End
            Begin VB.Label Label17 
               Caption         =   "Teléfono:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   21
               Top             =   1740
               Width           =   1905
            End
            Begin VB.Label Label22 
               Caption         =   "Celular:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   20
               Top             =   2070
               Width           =   1905
            End
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   405
            Left            =   4800
            TabIndex        =   32
            Top             =   120
            Width           =   6615
            _Version        =   65536
            _ExtentX        =   11668
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "DNI - 07521154 / IKEHARA PUNK MIGUEL ANGEL "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel16 
            Height          =   495
            Left            =   630
            TabIndex        =   33
            Top             =   60
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Referencias Personales"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            Picture         =   "AteCli_frm_005.frx":0A8E
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_IngSol_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_Fam_TipPar_Click()
   Call gs_SetFocus(txt_Fam_ApePat)
End Sub

Private Sub cmb_Fam_TipPar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Fam_TipPar_Click
   End If
End Sub

Private Sub cmb_NFa_TipPar_Click()
   Call gs_SetFocus(txt_NFa_ApePat)
End Sub

Private Sub cmb_NFa_TipPar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_NFa_TipPar_Click
   End If
End Sub

Private Sub cmd_Acepta_Click()
   If cmb_Fam_TipPar.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Parentesco de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Princi.Tab = 0
      Call gs_SetFocus(cmb_Fam_TipPar)
      Exit Sub
   End If
   
   If Len(Trim(txt_Fam_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Princi.Tab = 0
      Call gs_SetFocus(txt_Fam_ApePat)
      Exit Sub
   End If
   
   If Len(Trim(txt_Fam_ApeMat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Materno de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Princi.Tab = 0
      Call gs_SetFocus(txt_Fam_ApeMat)
      Exit Sub
   End If
   
   If Len(Trim(txt_Fam_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Princi.Tab = 0
      Call gs_SetFocus(txt_Fam_Nombre)
      Exit Sub
   End If
   
   If Len(Trim(txt_Fam_Telefo.Text)) = 0 And Len(Trim(txt_Fam_Celula.Text)) = 0 Then
      MsgBox "Debe ingresar algún Teléfono de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Princi.Tab = 0
      Call gs_SetFocus(txt_Fam_Telefo)
      Exit Sub
   End If
   
   
   If cmb_NFa_TipPar.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Parentesco de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Princi.Tab = 1
      Call gs_SetFocus(cmb_NFa_TipPar)
      Exit Sub
   End If
   
   If Len(Trim(txt_NFa_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Princi.Tab = 1
      Call gs_SetFocus(txt_NFa_ApePat)
      Exit Sub
   End If
   
   If Len(Trim(txt_NFa_ApeMat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Materno de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Princi.Tab = 1
      Call gs_SetFocus(txt_NFa_ApeMat)
      Exit Sub
   End If
   
   If Len(Trim(txt_NFa_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Princi.Tab = 1
      Call gs_SetFocus(txt_NFa_Nombre)
      Exit Sub
   End If
   
   If Len(Trim(txt_NFa_Telefo.Text)) = 0 And Len(Trim(txt_NFa_Celula.Text)) = 0 Then
      MsgBox "Debe ingresar algún Teléfono de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Princi.Tab = 1
      Call gs_SetFocus(txt_NFa_Telefo)
      Exit Sub
   End If
   
   Call modatecli_gs_Limpia_Refere(1)
   Call modatecli_gs_Limpia_Refere(2)
   
   modatecli_g_arr_Refere(1).Refere_TipPar = cmb_Fam_TipPar.ItemData(cmb_Fam_TipPar.ListIndex)
   modatecli_g_arr_Refere(1).Refere_ApePat = txt_Fam_ApePat.Text
   modatecli_g_arr_Refere(1).Refere_ApeMat = txt_Fam_ApeMat.Text
   modatecli_g_arr_Refere(1).Refere_Nombre = txt_Fam_Nombre.Text
   modatecli_g_arr_Refere(1).Refere_Telefo = txt_Fam_Telefo.Text
   modatecli_g_arr_Refere(1).Refere_Celula = txt_Fam_Celula.Text
   
   modatecli_g_arr_Refere(2).Refere_TipPar = cmb_Fam_TipPar.ItemData(cmb_NFa_TipPar.ListIndex)
   modatecli_g_arr_Refere(2).Refere_ApePat = txt_NFa_ApePat.Text
   modatecli_g_arr_Refere(2).Refere_ApeMat = txt_NFa_ApeMat.Text
   modatecli_g_arr_Refere(2).Refere_Nombre = txt_NFa_Nombre.Text
   modatecli_g_arr_Refere(2).Refere_Telefo = txt_NFa_Telefo.Text
   modatecli_g_arr_Refere(2).Refere_Celula = txt_NFa_Celula.Text
   
   modatecli_g_int_RefereTit = 2
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("Al salir de esta manera perderá la información ingresada. ¿Está seguro de salir de la ventana?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt & " Ingreso de Solicitud de Crédito"
   
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Limpia
   
   If modatecli_g_int_RefereTit = 2 Then
      Call gs_BuscarCombo_Item(cmb_Fam_TipPar, modatecli_g_arr_Refere(1).Refere_TipPar)
      txt_Fam_ApePat.Text = modatecli_g_arr_Refere(1).Refere_ApePat
      txt_Fam_ApeMat.Text = modatecli_g_arr_Refere(1).Refere_ApeMat
      txt_Fam_Nombre.Text = modatecli_g_arr_Refere(1).Refere_Nombre
      txt_Fam_Telefo.Text = modatecli_g_arr_Refere(1).Refere_Telefo
      txt_Fam_Celula.Text = modatecli_g_arr_Refere(1).Refere_Celula
   
   
      Call gs_BuscarCombo_Item(cmb_NFa_TipPar, modatecli_g_arr_Refere(2).Refere_TipPar)
      txt_NFa_ApePat.Text = modatecli_g_arr_Refere(2).Refere_ApePat
      txt_NFa_ApeMat.Text = modatecli_g_arr_Refere(2).Refere_ApeMat
      txt_NFa_Nombre.Text = modatecli_g_arr_Refere(2).Refere_Nombre
      txt_NFa_Telefo.Text = modatecli_g_arr_Refere(2).Refere_Telefo
      txt_NFa_Celula.Text = modatecli_g_arr_Refere(2).Refere_Celula
   End If
   
   Call gs_SetFocus(cmb_Fam_TipPar)
   
   tab_Princi.Tab = 0
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub txt_Fam_ApePat_GotFocus()
   Call gs_SelecTodo(txt_Fam_ApePat)
End Sub

Private Sub txt_Fam_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Fam_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_Fam_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_Fam_ApeMat)
End Sub

Private Sub txt_Fam_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Fam_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_Fam_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Fam_Nombre)
End Sub

Private Sub txt_Fam_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Fam_Telefo)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_Fam_Telefo_GotFocus()
   Call gs_SelecTodo(txt_Fam_Telefo)
End Sub

Private Sub txt_Fam_Telefo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Fam_Celula)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_()")
   End If
End Sub

Private Sub txt_Fam_Celula_GotFocus()
   Call gs_SelecTodo(txt_Fam_Celula)
End Sub

Private Sub txt_Fam_Celula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Acepta)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_()")
   End If
End Sub

Private Sub txt_NFa_ApePat_GotFocus()
   Call gs_SelecTodo(txt_NFa_ApePat)
End Sub

Private Sub txt_NFa_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NFa_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_NFa_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_NFa_ApeMat)
End Sub

Private Sub txt_NFa_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NFa_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_NFa_Nombre_GotFocus()
   Call gs_SelecTodo(txt_NFa_Nombre)
End Sub

Private Sub txt_NFa_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NFa_Telefo)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_NFa_Telefo_GotFocus()
   Call gs_SelecTodo(txt_NFa_Telefo)
End Sub

Private Sub txt_NFa_Telefo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NFa_Celula)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_()")
   End If
End Sub

Private Sub txt_NFa_Celula_GotFocus()
   Call gs_SelecTodo(txt_NFa_Celula)
End Sub

Private Sub txt_NFa_Celula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Acepta)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_()")
   End If
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_Fam_TipPar, 1, "212")
   Call moddat_gs_Carga_LisIte_Combo(cmb_NFa_TipPar, 1, "213")
End Sub

Private Sub fs_Limpia()
   cmb_Fam_TipPar.ListIndex = -1
   txt_Fam_ApePat.Text = ""
   txt_Fam_ApeMat.Text = ""
   txt_Fam_Nombre.Text = ""
   txt_Fam_Telefo.Text = ""
   txt_Fam_Celula.Text = ""

   cmb_NFa_TipPar.ListIndex = -1
   txt_NFa_ApePat.Text = ""
   txt_NFa_ApeMat.Text = ""
   txt_NFa_Nombre.Text = ""
   txt_NFa_Telefo.Text = ""
   txt_NFa_Celula.Text = ""
End Sub
