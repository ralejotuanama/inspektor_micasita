VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_MntCli_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3990
   ClientLeft      =   1500
   ClientTop       =   1740
   ClientWidth     =   11625
   Icon            =   "AteCli_frm_103.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   7011
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   1125
         Left            =   30
         TabIndex        =   10
         Top             =   2010
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
         Begin VB.ComboBox cmb_SegAct 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   765
         End
         Begin VB.ComboBox cmb_ActSec 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   750
            Width           =   9525
         End
         Begin VB.ComboBox cmb_ActEco 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   9525
         End
         Begin VB.Label Label11 
            Caption         =   "2da Actividad:"
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   420
            Width           =   1785
         End
         Begin VB.Label Label2 
            Caption         =   "Activ. Econ. Secundaria:"
            Height          =   315
            Left            =   90
            TabIndex        =   13
            Top             =   750
            Width           =   1845
         End
         Begin VB.Label Label19 
            Caption         =   "Activ. Econ. Principal:"
            Height          =   405
            Left            =   90
            TabIndex        =   11
            Top             =   90
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   5
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
            Height          =   495
            Left            =   630
            TabIndex        =   6
            Top             =   60
            Width           =   6465
            _Version        =   65536
            _ExtentX        =   11404
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes - Actividades Económicas"
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
            Picture         =   "AteCli_frm_103.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   7
         Top             =   750
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1950
            TabIndex        =   8
            Top             =   60
            Width           =   9525
            _Version        =   65536
            _ExtentX        =   16801
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1 - 07522154 / IKEHARA PUNK MIGUEL ANGEL"
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
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   9
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   735
         Left            =   30
         TabIndex        =   12
         Top             =   1230
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_103.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Datos del Crédito"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   735
         Left            =   30
         TabIndex        =   15
         Top             =   3180
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_SimCre 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_103.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_FlgFrm     As Integer

Private Sub cmb_ActEco_Click()
   If l_int_FlgFrm = 0 Then
      Exit Sub
   End If

   If cmb_ActEco.ListIndex > -1 Then
      moddat_g_int_FlgAct_1 = 1
   
      moddat_g_int_OrdAct = 1
      
      Select Case cmb_ActEco.ItemData(cmb_ActEco.ListIndex)
         Case 11: frm_MntCli_04.Show 1
         Case 21: frm_MntCli_05.Show 1
         Case 31: frm_MntCli_06.Show 1
         Case 41: frm_MntCli_07.Show 1
         Case 51: frm_MntCli_08.Show 1
         Case 61: frm_MntCli_10.Show 1
      End Select
   End If
End Sub

Private Sub cmb_ActSec_Click()
   If l_int_FlgFrm = 0 Then
      Exit Sub
   End If
   
   If cmb_ActSec.ListIndex > -1 Then
      moddat_g_int_FlgAct_1 = 1
   
      moddat_g_int_OrdAct = 2
      
      Select Case cmb_ActSec.ItemData(cmb_ActSec.ListIndex)
         Case 11: frm_MntCli_04.Show 1
         Case 21: frm_MntCli_05.Show 1
         Case 31: frm_MntCli_06.Show 1
         Case 41: frm_MntCli_07.Show 1
         Case 51: frm_MntCli_08.Show 1
         Case 61: frm_MntCli_10.Show 1
      End Select
   End If
End Sub

Private Sub cmb_ActEco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ActEco_Click
   End If
End Sub

Private Sub cmb_SegAct_Click()
   If cmb_SegAct.ListIndex = -1 Then
      cmb_ActSec.ListIndex = -1
      cmb_ActSec.Enabled = False
   Else
      If cmb_SegAct.ItemData(cmb_SegAct.ListIndex) = 1 Then
         cmb_ActSec.Enabled = True
         Call gs_SetFocus(cmb_ActSec)
      Else
         cmb_ActSec.ListIndex = -1
         cmb_ActSec.Enabled = False
         
         Call moddat_gs_Inicia_ActEco(moddat_g_int_TipCli, 2)
         
         Call gs_SetFocus(cmd_Salida)
      End If
   End If
End Sub

Private Sub cmd_Salida_Click()
   If moddat_g_int_FlgAct_1 = 2 Then
      If cmb_SegAct.ListIndex = -1 Then
         MsgBox "Debe seleccionar si el Cliente presenta Segunda Actividad Económica.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ActEco)
         Exit Sub
      End If
      
      If cmb_SegAct.ItemData(cmb_SegAct.ListIndex) = 1 Then
         If cmb_ActSec.ListIndex = -1 Then
            MsgBox "Debe ingresar la 2da. Actividad Económica..", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_ActSec)
            Exit Sub
         End If
         
         If moddat_g_int_TipCli = 1 Then
            If moddat_g_arr_ActEco_Tit(2).ActEco_TipAct = 0 Then
               MsgBox "Debe ingresar la 2da. Actividad Económica..", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(cmb_ActSec)
               Exit Sub
            End If
         Else
            If moddat_g_arr_ActEco_Cyg(2).ActEco_TipAct = 0 Then
               MsgBox "Debe ingresar la 2da. Actividad Económica..", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(cmb_ActSec)
               Exit Sub
            End If
         End If
      End If
   End If
   
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
   l_int_FlgFrm = 0

   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   moddat_g_int_FlgAct_1 = 1
   
   If moddat_g_int_TipCli = 1 Then
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
      
      If moddat_g_arr_ActEco_Tit(1).ActEco_TipAct > 0 Then
         Call gs_BuscarCombo_Item(cmb_ActEco, moddat_g_arr_ActEco_Tit(1).ActEco_TipAct)
         
         moddat_g_int_FlgAct_1 = 2
         
         If moddat_g_arr_ActEco_Tit(2).ActEco_TipAct > 0 Then
            Call gs_BuscarCombo_Item(cmb_SegAct, 1)
            Call gs_BuscarCombo_Item(cmb_ActSec, moddat_g_arr_ActEco_Tit(2).ActEco_TipAct)
         Else
            Call gs_BuscarCombo_Item(cmb_SegAct, 2)
            
            cmb_ActSec.ListIndex = -1
            cmb_ActSec.Enabled = False
         End If
      Else
         Call gs_SetFocus(cmb_ActEco)
      End If
   Else
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli & "(" & CStr(moddat_g_int_CygTDo) & " - " & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom & ")"
   
      If moddat_g_arr_ActEco_Cyg(1).ActEco_TipAct > 0 Then
         Call gs_BuscarCombo_Item(cmb_ActEco, moddat_g_arr_ActEco_Cyg(1).ActEco_TipAct)
         
         moddat_g_int_FlgAct_1 = 2
         
         If moddat_g_arr_ActEco_Cyg(2).ActEco_TipAct > 0 Then
            Call gs_BuscarCombo_Item(cmb_SegAct, 1)
            Call gs_BuscarCombo_Item(cmb_ActSec, moddat_g_arr_ActEco_Cyg(2).ActEco_TipAct)
         Else
            Call gs_BuscarCombo_Item(cmb_SegAct, 2)
            
            cmb_ActSec.ListIndex = -1
            cmb_ActSec.Enabled = False
         End If
      Else
         Call gs_SetFocus(cmb_ActEco)
      End If
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
   
   l_int_FlgFrm = 1
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_ActEco, 1, "008")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ActSec, 1, "008")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_SegAct, 1, "214")
End Sub

Private Sub fs_Limpia()
   cmb_ActEco.ListIndex = -1
   
   cmb_SegAct.ListIndex = -1
   cmb_ActSec.ListIndex = -1
   
   cmb_ActSec.Enabled = False
End Sub

