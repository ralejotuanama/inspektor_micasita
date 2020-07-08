VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Tra_CarAFP_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "AteCli_frm_582.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   4305
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11640
      _Version        =   65536
      _ExtentX        =   20532
      _ExtentY        =   7594
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
         Height          =   1125
         Left            =   30
         TabIndex        =   1
         Top             =   1920
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1984
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
            Height          =   675
            Left            =   1440
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Text            =   "AteCli_frm_582.frx":000C
            Top             =   390
            Width           =   10005
         End
         Begin Threed.SSPanel pnl_DesOcu 
            Height          =   315
            Left            =   1440
            TabIndex        =   3
            Top             =   60
            Width           =   10005
            _Version        =   65536
            _ExtentX        =   17648
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
         Begin VB.Label Label7 
            Caption         =   "Observación:"
            Height          =   495
            Left            =   60
            TabIndex        =   5
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label Label8 
            Caption         =   "Ocurrencia:"
            Height          =   315
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   1155
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   6
         Top             =   750
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "AteCli_frm_582.frx":0010
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_582.frx":0452
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Height          =   315
            Left            =   630
            TabIndex        =   10
            Top             =   30
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud de Crédito Hipotecario"
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   315
            Left            =   630
            TabIndex        =   11
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Descargo de Observaciones"
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
            Picture         =   "AteCli_frm_582.frx":0894
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   1155
         Left            =   30
         TabIndex        =   12
         Top             =   3090
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   2037
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
         Begin VB.TextBox txt_Descar 
            Height          =   1035
            Left            =   1440
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Text            =   "AteCli_frm_582.frx":0B9E
            Top             =   60
            Width           =   10035
         End
         Begin VB.Label Label10 
            Caption         =   "Descargo:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1035
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   435
         Left            =   30
         TabIndex        =   15
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   16
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
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
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   4440
            TabIndex        =   17
            Top             =   60
            Width           =   7035
            _Version        =   65536
            _ExtentX        =   12409
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
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   3690
            TabIndex        =   18
            Top             =   60
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "frm_Tra_CarAFP_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim l_int_CodIns        As Integer

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_Descar.Text)) = 0 Then
      MsgBox "Debe ingresar el Descargo de la Observación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descar)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_str_DesObs = txt_Descar.Text
   moddat_g_int_FlgAct_1 = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   moddat_g_int_CodIns = 33
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   txt_Descar.Text = ""
   
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT MAX(B.SEGDET_CODINS) AS CODINS FROM TRA_SEGDET B WHERE B.SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY B.SEGFECCRE DESC, B.SEGHORCRE DESC "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
'      Exit Sub
'   End If
'
'   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
'      g_rst_Genera.MoveFirst
'      moddat_g_int_CodIns = g_rst_Genera!CODINS
'      l_int_CodIns = g_rst_Genera!CODINS
'   End If
'
'   g_rst_Genera.Close
'   Set g_rst_Genera = Nothing
'
'   If moddat_g_int_CodIns = 32 Then l_int_CodIns = 31
'   If moddat_g_int_CodIns = 42 Then l_int_CodIns = 41
'
   'Buscando Observación Pendiente
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(moddat_g_int_CodIns) & " AND " 'l_int_CodIns
   g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 99 AND "
   g_str_Parame = g_str_Parame & "SEGDET_NUMOBS = " & CStr(moddat_g_int_NumObs)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
     pnl_DesOcu.Caption = "Día: " & gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE)) & " - " & gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000")) & " - " & moddat_gf_Consulta_ParDes("004", Format(g_rst_Princi!SEGDET_CODOCU, "000000"))
     
     txt_Observ.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub txt_Descar_GotFocus()
   Call gs_SelecTodo(txt_Descar)
End Sub

Private Sub txt_Descar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub


