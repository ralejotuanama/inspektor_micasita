VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm_Imprim_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   7755
   ClientLeft      =   540
   ClientTop       =   2265
   ClientWidth     =   13695
   Icon            =   "com_frm_001.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   13695
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      _Version        =   65536
      _ExtentX        =   24156
      _ExtentY        =   13679
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
         Height          =   6135
         Left            =   30
         TabIndex        =   1
         Top             =   750
         Width           =   13605
         _Version        =   65536
         _ExtentX        =   23998
         _ExtentY        =   10821
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
         Begin VB.TextBox txt_Imprim 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6015
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   2
            Text            =   "com_frm_001.frx":000C
            Top             =   60
            Width           =   13485
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   13605
         _Version        =   65536
         _ExtentX        =   23998
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
            Height          =   480
            Left            =   630
            TabIndex        =   4
            Top             =   90
            Width           =   5175
            _Version        =   65536
            _ExtentX        =   9128
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Impresión de Datos"
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
            Left            =   90
            Picture         =   "com_frm_001.frx":0012
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   30
         TabIndex        =   5
         Top             =   6930
         Width           =   13605
         _Version        =   65536
         _ExtentX        =   23998
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12900
            Picture         =   "com_frm_001.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   11520
            Picture         =   "com_frm_001.frx":0896
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_ArcTxt 
            Height          =   675
            Left            =   12210
            Picture         =   "com_frm_001.frx":0CD8
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   30
            Width           =   675
         End
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   420
            Top             =   210
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
End
Attribute VB_Name = "frm_Imprim_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ArcTxt_Click()
   Dim r_int_NumFil     As Integer
   Dim r_int_Contad     As Integer
   
   On Error GoTo cmd_ArcTxt_Error
   
   If MsgBox("¿Está seguro de Guardar el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   dlg_Guarda.Filter = "Texto (*.txt)|*.txt"
   dlg_Guarda.ShowSave
   
   
   'Crear Archivo
   r_int_NumFil = FreeFile
   Open dlg_Guarda.FileName For Output As r_int_NumFil
                
   For r_int_Contad = 1 To UBound(g_arr_Imprim)
      If g_arr_Imprim(r_int_Contad).Imprim_ConLen = "SP" Then
         Print #r_int_NumFil, Chr(12)
      Else
         If Len(Trim(g_arr_Imprim(r_int_Contad).Imprim_ConLen)) > 0 Then
            Print #r_int_NumFil, g_arr_Imprim(r_int_Contad).Imprim_ConLen
         Else
            Print #r_int_NumFil, ""
         End If
      End If
   Next r_int_Contad
                
   'Cerrando Archivo
   Close #r_int_NumFil
   
cmd_ArcTxt_Error:
   Exit Sub
   
End Sub

Private Sub cmd_Imprim_Click()
   On Error GoTo cmd_Imprim_Error
   
   If MsgBox("¿Está seguro de Imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   dlg_Guarda.CancelError = True
   dlg_Guarda.ShowPrinter
   
   Call gs_Imprim(1, 8, 1)
   
cmd_Imprim_Error:
   Exit Sub
   
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_int_Contad     As Integer
   
   Screen.MousePointer = 11
   
   Call SendMessage(txt_Imprim.hWnd, EM_LIMITTEXT, 0, ByVal 0&)
   
   Me.Caption = modgen_g_str_NomPlt

   txt_Imprim.Locked = True
   txt_Imprim.Text = ""
   
   For r_int_Contad = 1 To UBound(g_arr_Imprim)
      If g_arr_Imprim(r_int_Contad).Imprim_ConLen = "SP" Then
         txt_Imprim.Text = txt_Imprim.Text & Chr(12)
      Else
         If Len(Trim(g_arr_Imprim(r_int_Contad).Imprim_ConLen)) > 0 Then
            txt_Imprim.Text = txt_Imprim.Text & Mid(g_arr_Imprim(r_int_Contad).Imprim_ConLen & Chr(13) & Chr(10), 6)
         Else
            txt_Imprim.Text = txt_Imprim.Text & Chr(13) & Chr(10)
         End If
      End If
   Next r_int_Contad
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub


