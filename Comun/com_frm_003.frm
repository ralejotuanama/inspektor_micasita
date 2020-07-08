VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_EnvMai_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3360
   ClientLeft      =   4755
   ClientTop       =   3765
   ClientWidth     =   5355
   Icon            =   "com_frm_003.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3345
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5355
      _Version        =   65536
      _ExtentX        =   9446
      _ExtentY        =   5900
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   735
         Left            =   30
         TabIndex        =   15
         Top             =   2550
         Width           =   5265
         _Version        =   65536
         _ExtentX        =   9287
         _ExtentY        =   1296
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
         Begin VB.CommandButton cmd_Enviar 
            Height          =   675
            Left            =   4560
            Picture         =   "com_frm_003.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Enviar Correo Electrónico"
            Top             =   30
            Width           =   675
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   690
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   120
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1755
         Left            =   30
         TabIndex        =   7
         Top             =   750
         Width           =   5265
         _Version        =   65536
         _ExtentX        =   9287
         _ExtentY        =   3096
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
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Index           =   0
            Left            =   1860
            MaxLength       =   120
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Index           =   1
            Left            =   1860
            MaxLength       =   120
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Index           =   2
            Left            =   1860
            MaxLength       =   120
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Index           =   3
            Left            =   1860
            MaxLength       =   120
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Index           =   4
            Left            =   1860
            MaxLength       =   120
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.Label Label1 
            Caption         =   "Dirección de Correo (1):"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1725
         End
         Begin VB.Label Label2 
            Caption         =   "Dirección de Correo (2):"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   390
            Width           =   1725
         End
         Begin VB.Label Label3 
            Caption         =   "Dirección de Correo (3):"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label4 
            Caption         =   "Dirección de Correo (4):"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   1050
            Width           =   1725
         End
         Begin VB.Label Label5 
            Caption         =   "Dirección de Correo (5):"
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   1380
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   5265
         _Version        =   65536
         _ExtentX        =   9287
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
            Left            =   600
            TabIndex        =   14
            Top             =   60
            Width           =   4000
            _Version        =   65536
            _ExtentX        =   7056
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Envío de E-Mail"
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
            Picture         =   "com_frm_003.frx":0316
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_EnvMai_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Enviar_Click()
   Dim r_int_Contad  As Integer
   Dim r_int_Posici  As Integer

   On Error GoTo Error_cmd_Enviar

   'Inicializa
   mps_Sesion.DownLoadMail = False
   mps_Sesion.SignOn
   mps_Mensaj.SessionID = mps_Sesion.SessionID
  
   'Envío
   mps_Mensaj.Compose
  
   r_int_Posici = 0
   For r_int_Contad = 0 To 4
      If Len(Trim(txt_DirEle(r_int_Contad).Text)) > 0 Then
         mps_Mensaj.RecipIndex = r_int_Posici
         mps_Mensaj.RecipDisplayName = txt_DirEle(r_int_Contad).Text
         
         r_int_Posici = r_int_Posici + 1
      Else
         Exit For
      End If
   Next r_int_Contad

   mps_Mensaj.MsgSubject = modgen_g_str_Mail_Asunto
   mps_Mensaj.MsgNoteText = modgen_g_str_Mail_Mensaj
   
   mps_Mensaj.Send
   DoEvents
  
  'Cierra la sesión
  mps_Sesion.SignOff
  Unload Me
  Exit Sub
  
Error_cmd_Enviar:
   MsgBox "Reintente nuevamente el envio del mail", vbExclamation, modgen_g_str_NomPlt
   
End Sub

Private Sub Form_Load()
   Dim r_int_Contad  As Integer
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   For r_int_Contad = 0 To 4
      txt_DirEle(r_int_Contad).Text = ""
   Next r_int_Contad
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub txt_DirEle_GotFocus(Index As Integer)
   Call gs_SelecTodo(txt_DirEle(Index))
End Sub

Private Sub txt_DirEle_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index = 4 Then
         Call gs_SetFocus(cmd_Enviar)
      Else
         Call gs_SetFocus(txt_DirEle(Index + 1))
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_@.")
   End If
End Sub

