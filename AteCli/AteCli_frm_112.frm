VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_SolCre_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   1905
   ClientTop       =   630
   ClientWidth     =   11625
   Icon            =   "AteCli_frm_112.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8895
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   15690
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
         Height          =   3255
         Left            =   30
         TabIndex        =   44
         Top             =   4770
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5741
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
         Begin VB.TextBox txt_NFa_ApePat_1 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   2220
            Width           =   3315
         End
         Begin VB.ComboBox cmb_NFa_TipPar_1 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1890
            Width           =   3315
         End
         Begin VB.TextBox txt_NFa_ApeMat_1 
            Height          =   315
            Left            =   8130
            MaxLength       =   30
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   2220
            Width           =   3315
         End
         Begin VB.TextBox txt_NFa_Nombre_1 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   2550
            Width           =   3315
         End
         Begin VB.TextBox txt_NFa_Telefo_1 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   2880
            Width           =   3315
         End
         Begin VB.TextBox txt_NFa_Celula_1 
            Height          =   315
            Left            =   8130
            MaxLength       =   12
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   2880
            Width           =   3315
         End
         Begin VB.TextBox txt_NFa_ApePat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   750
            Width           =   3315
         End
         Begin VB.ComboBox cmb_NFa_TipPar 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   420
            Width           =   3315
         End
         Begin VB.TextBox txt_NFa_ApeMat 
            Height          =   315
            Left            =   8130
            MaxLength       =   30
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   750
            Width           =   3315
         End
         Begin VB.TextBox txt_NFa_Nombre 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   1080
            Width           =   3315
         End
         Begin VB.TextBox txt_NFa_Telefo 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   1410
            Width           =   3315
         End
         Begin VB.TextBox txt_NFa_Celula 
            Height          =   315
            Left            =   8130
            MaxLength       =   12
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   1410
            Width           =   3315
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   60
            Left            =   30
            TabIndex        =   58
            Top             =   1770
            Width           =   11355
            _Version        =   65536
            _ExtentX        =   20029
            _ExtentY        =   106
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
         End
         Begin VB.Label Label28 
            Caption         =   "Tipo Parentesco:"
            Height          =   315
            Left            =   90
            TabIndex        =   64
            Top             =   1890
            Width           =   1635
         End
         Begin VB.Label Label27 
            Caption         =   "Apellido Paterno:"
            Height          =   315
            Left            =   90
            TabIndex        =   63
            Top             =   2220
            Width           =   1905
         End
         Begin VB.Label Label26 
            Caption         =   "Apellido Materno:"
            Height          =   315
            Left            =   6210
            TabIndex        =   62
            Top             =   2220
            Width           =   1635
         End
         Begin VB.Label Label25 
            Caption         =   "Nombres:"
            Height          =   315
            Left            =   90
            TabIndex        =   61
            Top             =   2550
            Width           =   1905
         End
         Begin VB.Label Label24 
            Caption         =   "Teléfono:"
            Height          =   315
            Left            =   90
            TabIndex        =   60
            Top             =   2880
            Width           =   1905
         End
         Begin VB.Label Label11 
            Caption         =   "Celular:"
            Height          =   315
            Left            =   6210
            TabIndex        =   59
            Top             =   2880
            Width           =   1065
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo Parentesco:"
            Height          =   315
            Left            =   90
            TabIndex        =   51
            Top             =   420
            Width           =   1905
         End
         Begin VB.Label Label9 
            Caption         =   "Apellido Paterno:"
            Height          =   315
            Left            =   90
            TabIndex        =   50
            Top             =   750
            Width           =   1905
         End
         Begin VB.Label Label8 
            Caption         =   "Apellido Materno:"
            Height          =   315
            Left            =   6210
            TabIndex        =   49
            Top             =   750
            Width           =   1635
         End
         Begin VB.Label Label7 
            Caption         =   "Nombres:"
            Height          =   315
            Left            =   90
            TabIndex        =   48
            Top             =   1080
            Width           =   1905
         End
         Begin VB.Label Label5 
            Caption         =   "Teléfono:"
            Height          =   315
            Left            =   90
            TabIndex        =   47
            Top             =   1410
            Width           =   1905
         End
         Begin VB.Label Label4 
            Caption         =   "Celular:"
            Height          =   315
            Left            =   6210
            TabIndex        =   46
            Top             =   1410
            Width           =   1065
         End
         Begin VB.Label Label6 
            Caption         =   "Referencias No Familiares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   45
            Top             =   90
            Width           =   2655
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   34
         Top             =   8070
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_SimCre 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_112.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_112.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Acepta 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_112.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Aceptar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   3225
         Left            =   30
         TabIndex        =   35
         Top             =   1500
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5689
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
         Begin VB.TextBox txt_Fam_ApePat_1 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   2190
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Fam_TipPar_1 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1860
            Width           =   3315
         End
         Begin VB.TextBox txt_Fam_ApeMat_1 
            Height          =   315
            Left            =   8130
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   2190
            Width           =   3315
         End
         Begin VB.TextBox txt_Fam_Nombre_1 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   2520
            Width           =   3315
         End
         Begin VB.TextBox txt_Fam_Telefo_1 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   2850
            Width           =   3315
         End
         Begin VB.TextBox txt_Fam_Celula_1 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8130
            MaxLength       =   12
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   2850
            Width           =   3315
         End
         Begin VB.TextBox txt_Fam_Celula 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8130
            MaxLength       =   12
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_Fam_Telefo 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_Fam_Nombre 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_Fam_ApeMat 
            Height          =   315
            Left            =   8130
            MaxLength       =   30
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Fam_TipPar 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_Fam_ApePat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   60
            Left            =   60
            TabIndex        =   43
            Top             =   1740
            Width           =   11355
            _Version        =   65536
            _ExtentX        =   20029
            _ExtentY        =   106
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
         End
         Begin VB.Label Label23 
            Caption         =   "Tipo Parentesco:"
            Height          =   315
            Left            =   90
            TabIndex        =   57
            Top             =   1860
            Width           =   1605
         End
         Begin VB.Label Label18 
            Caption         =   "Apellido Paterno:"
            Height          =   315
            Left            =   90
            TabIndex        =   56
            Top             =   2190
            Width           =   1905
         End
         Begin VB.Label Label16 
            Caption         =   "Apellido Materno:"
            Height          =   315
            Left            =   6210
            TabIndex        =   55
            Top             =   2190
            Width           =   1905
         End
         Begin VB.Label Label15 
            Caption         =   "Nombres:"
            Height          =   315
            Left            =   90
            TabIndex        =   54
            Top             =   2520
            Width           =   1905
         End
         Begin VB.Label Label13 
            Caption         =   "Teléfono:"
            Height          =   315
            Left            =   90
            TabIndex        =   53
            Top             =   2850
            Width           =   1905
         End
         Begin VB.Label Label12 
            Caption         =   "Celular:"
            Height          =   315
            Left            =   6210
            TabIndex        =   52
            Top             =   2850
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "Referencias Familiares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   42
            Top             =   90
            Width           =   1965
         End
         Begin VB.Label Label22 
            Caption         =   "Celular:"
            Height          =   315
            Left            =   6210
            TabIndex        =   41
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label17 
            Caption         =   "Teléfono:"
            Height          =   315
            Left            =   90
            TabIndex        =   40
            Top             =   1380
            Width           =   1515
         End
         Begin VB.Label Label14 
            Caption         =   "Nombres:"
            Height          =   315
            Left            =   90
            TabIndex        =   39
            Top             =   1050
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Materno:"
            Height          =   315
            Left            =   6210
            TabIndex        =   38
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label Label2 
            Caption         =   "Apellido Paterno:"
            Height          =   315
            Left            =   90
            TabIndex        =   37
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo Parentesco:"
            Height          =   315
            Left            =   90
            TabIndex        =   36
            Top             =   390
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   615
         Left            =   30
         TabIndex        =   27
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1085
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   495
            Left            =   660
            TabIndex        =   28
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito - Referencias Personales"
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
            Picture         =   "AteCli_frm_112.frx":0A62
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   29
         Top             =   690
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1890
            TabIndex        =   30
            Top             =   60
            Width           =   9585
            _Version        =   65536
            _ExtentX        =   16907
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1890
            TabIndex        =   31
            Top             =   390
            Width           =   9585
            _Version        =   65536
            _ExtentX        =   16907
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   33
            Top             =   390
            Width           =   1755
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   32
            Top             =   60
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_SolCre_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_Fam_TipPar_1_Click()
   Call gs_SetFocus(txt_Fam_ApePat_1)
End Sub

Private Sub cmb_Fam_TipPar_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Fam_TipPar_1_Click
   End If
End Sub

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
      Call gs_SetFocus(cmb_Fam_TipPar)
      Exit Sub
   End If
   
   If cmb_Fam_TipPar.ItemData(cmb_Fam_TipPar.ListIndex) = 8 Then
      MsgBox "La Primera Referencia Familiar no puede ser << NINGUNO >>.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Fam_TipPar)
      Exit Sub
   End If
   
   If Len(Trim(txt_Fam_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Fam_ApePat)
      Exit Sub
   End If
   
   'If Len(Trim(txt_Fam_ApeMat.Text)) = 0 Then
   '   MsgBox "Debe ingresar el Apellido Materno de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(txt_Fam_ApeMat)
   '   Exit Sub
   'End If
   
   If Len(Trim(txt_Fam_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Fam_Nombre)
      Exit Sub
   End If
   
   If Len(Trim(txt_Fam_Telefo.Text)) = 0 Then
      MsgBox "Debe ingresar el Teléfono de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Fam_Telefo)
      Exit Sub
   End If
   
   If cmb_Fam_TipPar_1.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Parentesco de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Fam_TipPar_1)
      Exit Sub
   End If
   
   If cmb_Fam_TipPar_1.ItemData(cmb_Fam_TipPar_1.ListIndex) <> 8 Then
      If Len(Trim(txt_Fam_ApePat_1.Text)) = 0 Then
         MsgBox "Debe ingresar el Apellido Paterno de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Fam_ApePat_1)
         Exit Sub
      End If
   
      If Len(Trim(txt_Fam_Nombre_1.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Fam_Nombre_1)
         Exit Sub
      End If
   
      If Len(Trim(txt_Fam_Telefo_1.Text)) = 0 Then
         MsgBox "Debe ingresar el Teléfono de la Referencia Familiar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Fam_Telefo_1)
         Exit Sub
      End If
   End If
   
   If cmb_NFa_TipPar.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Parentesco de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NFa_TipPar)
      Exit Sub
   End If
   
   If cmb_NFa_TipPar.ItemData(cmb_NFa_TipPar.ListIndex) = 8 Then
      MsgBox "La Primera Referencia No Familiar no puede ser << NINGUNO >>.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NFa_TipPar)
      Exit Sub
   End If
   
   If Len(Trim(txt_NFa_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NFa_ApePat)
      Exit Sub
   End If
   
   If Len(Trim(txt_NFa_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NFa_Nombre)
      Exit Sub
   End If
   
   If Len(Trim(txt_NFa_Telefo.Text)) = 0 Then
      MsgBox "Debe ingresar el Teléfono de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NFa_Telefo)
      Exit Sub
   End If
   
   If cmb_NFa_TipPar_1.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Parentesco de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NFa_TipPar_1)
      Exit Sub
   End If
   
   If cmb_NFa_TipPar_1.ItemData(cmb_NFa_TipPar_1.ListIndex) <> 8 Then
      If Len(Trim(txt_NFa_ApePat_1.Text)) = 0 Then
         MsgBox "Debe ingresar el Apellido Paterno de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NFa_ApePat_1)
         Exit Sub
      End If
      
      If Len(Trim(txt_NFa_Nombre_1.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NFa_Nombre_1)
         Exit Sub
      End If
      
      If Len(Trim(txt_NFa_Telefo_1.Text)) = 0 Then
         MsgBox "Debe ingresar el Teléfono de la Referencia No Familiar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NFa_Telefo_1)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call modatecli_gs_Limpia_Refere(1)
   Call modatecli_gs_Limpia_Refere(2)
   Call modatecli_gs_Limpia_Refere(3)
   Call modatecli_gs_Limpia_Refere(4)
   
   modatecli_g_arr_Refere(1).Refere_TipPar = cmb_Fam_TipPar.ItemData(cmb_Fam_TipPar.ListIndex)
   modatecli_g_arr_Refere(1).Refere_ApePat = txt_Fam_ApePat.Text
   modatecli_g_arr_Refere(1).Refere_ApeMat = txt_Fam_ApeMat.Text
   modatecli_g_arr_Refere(1).Refere_Nombre = txt_Fam_Nombre.Text
   modatecli_g_arr_Refere(1).Refere_Telefo = txt_Fam_Telefo.Text
   modatecli_g_arr_Refere(1).Refere_Celula = txt_Fam_Celula.Text
   
   modatecli_g_arr_Refere(2).Refere_TipPar = cmb_Fam_TipPar_1.ItemData(cmb_Fam_TipPar_1.ListIndex)
   
   If cmb_Fam_TipPar_1.ItemData(cmb_Fam_TipPar_1.ListIndex) <> 8 Then
      modatecli_g_arr_Refere(2).Refere_ApePat = txt_Fam_ApePat_1.Text
      modatecli_g_arr_Refere(2).Refere_ApeMat = txt_Fam_ApeMat_1.Text
      modatecli_g_arr_Refere(2).Refere_Nombre = txt_Fam_Nombre_1.Text
      modatecli_g_arr_Refere(2).Refere_Telefo = txt_Fam_Telefo_1.Text
      modatecli_g_arr_Refere(2).Refere_Celula = txt_Fam_Celula_1.Text
   Else
      modatecli_g_arr_Refere(2).Refere_ApePat = ""
      modatecli_g_arr_Refere(2).Refere_ApeMat = ""
      modatecli_g_arr_Refere(2).Refere_Nombre = ""
      modatecli_g_arr_Refere(2).Refere_Telefo = ""
      modatecli_g_arr_Refere(2).Refere_Celula = ""
   End If
   
   modatecli_g_arr_Refere(3).Refere_TipPar = cmb_NFa_TipPar.ItemData(cmb_NFa_TipPar.ListIndex)
   modatecli_g_arr_Refere(3).Refere_ApePat = txt_NFa_ApePat.Text
   modatecli_g_arr_Refere(3).Refere_ApeMat = txt_NFa_ApeMat.Text
   modatecli_g_arr_Refere(3).Refere_Nombre = txt_NFa_Nombre.Text
   modatecli_g_arr_Refere(3).Refere_Telefo = txt_NFa_Telefo.Text
   modatecli_g_arr_Refere(3).Refere_Celula = txt_NFa_Celula.Text
   
   modatecli_g_arr_Refere(4).Refere_TipPar = cmb_NFa_TipPar_1.ItemData(cmb_NFa_TipPar_1.ListIndex)
   
   If cmb_NFa_TipPar_1.ItemData(cmb_NFa_TipPar_1.ListIndex) <> 8 Then
      modatecli_g_arr_Refere(4).Refere_ApePat = txt_NFa_ApePat_1.Text
      modatecli_g_arr_Refere(4).Refere_ApeMat = txt_NFa_ApeMat_1.Text
      modatecli_g_arr_Refere(4).Refere_Nombre = txt_NFa_Nombre_1.Text
      modatecli_g_arr_Refere(4).Refere_Telefo = txt_NFa_Telefo_1.Text
      modatecli_g_arr_Refere(4).Refere_Celula = txt_NFa_Celula_1.Text
   Else
      modatecli_g_arr_Refere(4).Refere_ApePat = ""
      modatecli_g_arr_Refere(4).Refere_ApeMat = ""
      modatecli_g_arr_Refere(4).Refere_Nombre = ""
      modatecli_g_arr_Refere(4).Refere_Telefo = ""
      modatecli_g_arr_Refere(4).Refere_Celula = ""
   End If
   
   modatecli_g_int_RefereTit = 2
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_11.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_Produc.Caption = moddat_gf_Consulta_Produc(moddat_g_str_CodPrd)
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
   
      Call gs_BuscarCombo_Item(cmb_Fam_TipPar_1, modatecli_g_arr_Refere(2).Refere_TipPar)
      txt_Fam_ApePat_1.Text = modatecli_g_arr_Refere(2).Refere_ApePat
      txt_Fam_ApeMat_1.Text = modatecli_g_arr_Refere(2).Refere_ApeMat
      txt_Fam_Nombre_1.Text = modatecli_g_arr_Refere(2).Refere_Nombre
      txt_Fam_Telefo_1.Text = modatecli_g_arr_Refere(2).Refere_Telefo
      txt_Fam_Celula_1.Text = modatecli_g_arr_Refere(2).Refere_Celula
   
   
      Call gs_BuscarCombo_Item(cmb_NFa_TipPar, modatecli_g_arr_Refere(3).Refere_TipPar)
      txt_NFa_ApePat.Text = modatecli_g_arr_Refere(3).Refere_ApePat
      txt_NFa_ApeMat.Text = modatecli_g_arr_Refere(3).Refere_ApeMat
      txt_NFa_Nombre.Text = modatecli_g_arr_Refere(3).Refere_Nombre
      txt_NFa_Telefo.Text = modatecli_g_arr_Refere(3).Refere_Telefo
      txt_NFa_Celula.Text = modatecli_g_arr_Refere(3).Refere_Celula
   
      Call gs_BuscarCombo_Item(cmb_NFa_TipPar_1, modatecli_g_arr_Refere(4).Refere_TipPar)
      txt_NFa_ApePat_1.Text = modatecli_g_arr_Refere(4).Refere_ApePat
      txt_NFa_ApeMat_1.Text = modatecli_g_arr_Refere(4).Refere_ApeMat
      txt_NFa_Nombre_1.Text = modatecli_g_arr_Refere(4).Refere_Nombre
      txt_NFa_Telefo_1.Text = modatecli_g_arr_Refere(4).Refere_Telefo
      txt_NFa_Celula_1.Text = modatecli_g_arr_Refere(4).Refere_Celula
   End If
   
   Call gs_SetFocus(cmb_Fam_TipPar)

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
      Call gs_SetFocus(cmb_Fam_TipPar_1)
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
      Call gs_SetFocus(cmb_NFa_TipPar_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_()")
   End If
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_Fam_TipPar, 1, "212")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Fam_TipPar_1, 1, "212")
   Call moddat_gs_Carga_LisIte_Combo(cmb_NFa_TipPar, 1, "213")
   Call moddat_gs_Carga_LisIte_Combo(cmb_NFa_TipPar_1, 1, "213")
End Sub

Private Sub fs_Limpia()
   cmb_Fam_TipPar.ListIndex = -1
   txt_Fam_ApePat.Text = ""
   txt_Fam_ApeMat.Text = ""
   txt_Fam_Nombre.Text = ""
   txt_Fam_Telefo.Text = ""
   txt_Fam_Celula.Text = ""

   cmb_Fam_TipPar_1.ListIndex = -1
   txt_Fam_ApePat_1.Text = ""
   txt_Fam_ApeMat_1.Text = ""
   txt_Fam_Nombre_1.Text = ""
   txt_Fam_Telefo_1.Text = ""
   txt_Fam_Celula_1.Text = ""

   cmb_NFa_TipPar.ListIndex = -1
   txt_NFa_ApePat.Text = ""
   txt_NFa_ApeMat.Text = ""
   txt_NFa_Nombre.Text = ""
   txt_NFa_Telefo.Text = ""
   txt_NFa_Celula.Text = ""

   cmb_NFa_TipPar_1.ListIndex = -1
   txt_NFa_ApePat_1.Text = ""
   txt_NFa_ApeMat_1.Text = ""
   txt_NFa_Nombre_1.Text = ""
   txt_NFa_Telefo_1.Text = ""
   txt_NFa_Celula_1.Text = ""
End Sub

Private Sub txt_Fam_ApePat_1_GotFocus()
   Call gs_SelecTodo(txt_Fam_ApePat_1)
End Sub

Private Sub txt_Fam_ApePat_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Fam_ApeMat_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_Fam_ApeMat_1_GotFocus()
   Call gs_SelecTodo(txt_Fam_ApeMat_1)
End Sub

Private Sub txt_Fam_ApeMat_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Fam_Nombre_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_Fam_Nombre_1_GotFocus()
   Call gs_SelecTodo(txt_Fam_Nombre_1)
End Sub

Private Sub txt_Fam_Nombre_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Fam_Telefo_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_Fam_Telefo_1_GotFocus()
   Call gs_SelecTodo(txt_Fam_Telefo_1)
End Sub

Private Sub txt_Fam_Telefo_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Fam_Celula_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_()")
   End If
End Sub

Private Sub txt_Fam_Celula_1_GotFocus()
   Call gs_SelecTodo(txt_Fam_Celula_1)
End Sub

Private Sub txt_Fam_Celula_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_NFa_TipPar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_()")
   End If
End Sub

Private Sub cmb_NFa_TipPar_1_Click()
   Call gs_SetFocus(txt_NFa_ApePat_1)
End Sub

Private Sub cmb_NFa_TipPar_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_NFa_TipPar_1_Click
   End If
End Sub

Private Sub txt_NFa_ApePat_1_GotFocus()
   Call gs_SelecTodo(txt_NFa_ApePat_1)
End Sub

Private Sub txt_NFa_ApePat_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NFa_ApeMat_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_NFa_ApeMat_1_GotFocus()
   Call gs_SelecTodo(txt_NFa_ApeMat_1)
End Sub

Private Sub txt_NFa_ApeMat_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NFa_Nombre_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_NFa_Nombre_1_GotFocus()
   Call gs_SelecTodo(txt_NFa_Nombre_1)
End Sub

Private Sub txt_NFa_Nombre_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NFa_Telefo_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .'")
   End If
End Sub

Private Sub txt_NFa_Telefo_1_GotFocus()
   Call gs_SelecTodo(txt_NFa_Telefo_1)
End Sub

Private Sub txt_NFa_Telefo_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NFa_Celula_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_()")
   End If
End Sub

Private Sub txt_NFa_Celula_1_GotFocus()
   Call gs_SelecTodo(txt_NFa_Celula_1)
End Sub

Private Sub txt_NFa_Celula_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Acepta)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_()")
   End If
End Sub


