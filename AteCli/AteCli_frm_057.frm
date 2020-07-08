VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Rechaz_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   2055
   ClientTop       =   2385
   ClientWidth     =   10905
   Icon            =   "AteCli_frm_057.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4515
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10905
      _Version        =   65536
      _ExtentX        =   19235
      _ExtentY        =   7964
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
         Height          =   765
         Left            =   30
         TabIndex        =   8
         Top             =   3690
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   9450
            Picture         =   "AteCli_frm_057.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10110
            Picture         =   "AteCli_frm_057.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   585
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   465
            Left            =   630
            TabIndex        =   6
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   820
            _StockProps     =   15
            Caption         =   "Rechazo de Solicitud"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   405
            Left            =   3960
            TabIndex        =   7
            Top             =   90
            Width           =   6825
            _Version        =   65536
            _ExtentX        =   12039
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   90
            Picture         =   "AteCli_frm_057.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   9
         Top             =   660
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
         _ExtentY        =   2514
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   10
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   556
            _StockProps     =   15
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
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   7170
            TabIndex        =   11
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_EjeVta 
            Height          =   315
            Left            =   1620
            TabIndex        =   12
            Top             =   1050
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Modali 
            Height          =   315
            Left            =   1620
            TabIndex        =   13
            Top             =   720
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1620
            TabIndex        =   14
            Top             =   390
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   7170
            TabIndex        =   15
            Top             =   60
            Width           =   3585
            _Version        =   65536
            _ExtentX        =   6324
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Situac 
            Height          =   315
            Left            =   3360
            TabIndex        =   16
            Top             =   60
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   556
            _StockProps     =   15
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
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_InsAct 
            Height          =   315
            Left            =   7170
            TabIndex        =   17
            Top             =   720
            Width           =   3585
            _Version        =   65536
            _ExtentX        =   6324
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label17 
            Caption         =   "Instancia Actual:"
            Height          =   315
            Left            =   5820
            TabIndex        =   24
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "F. Ingreso Solic.:"
            Height          =   315
            Left            =   5820
            TabIndex        =   22
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label3 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label24 
            Caption         =   "Moneda Prést.:"
            Height          =   315
            Left            =   5820
            TabIndex        =   18
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1515
         Left            =   30
         TabIndex        =   25
         Top             =   2130
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
         _ExtentY        =   2672
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
         Begin VB.TextBox txt_Observ 
            Height          =   1095
            Left            =   1620
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Text            =   "AteCli_frm_057.frx":0CD2
            Top             =   390
            Width           =   9135
         End
         Begin VB.ComboBox cmb_MotRec 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   9135
         End
         Begin VB.Label Label23 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   60
            TabIndex        =   27
            Top             =   390
            Width           =   1245
         End
         Begin VB.Label Label8 
            Caption         =   "Motivo Rechazo:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "frm_Rechaz_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_MotRec_Click()
   Call gs_SetFocus(txt_Observ)
End Sub

Private Sub cmb_MotRec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MotRec_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_MotRec.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Motivo del Rechazo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MotRec)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de rechazar la Solicitud?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   moddat_g_int_MotRec = cmb_MotRec.ItemData(cmb_MotRec.ListIndex)
   moddat_g_str_Observ = txt_Observ.Text
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod
   pnl_EjeVta.Caption = moddat_g_str_EjeVta
   pnl_Moneda.Caption = moddat_g_str_Moneda
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_InsAct.Caption = moddat_gf_Consulta_Pardes("002", Trim(moddat_g_int_InsAct))

   Call moddat_gs_Carga_MotRec(cmb_MotRec, moddat_g_int_InsAct)
   
   cmb_MotRec.ListIndex = -1
   txt_Observ.Text = ""
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub


