VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_10 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   1905
   ClientTop       =   3960
   ClientWidth     =   11700
   Icon            =   "AteCli_frm_130.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3915
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   6906
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
         Height          =   1815
         Left            =   30
         TabIndex        =   7
         Top             =   1230
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   3201
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
         Begin VB.ComboBox cmb_CodCiu 
            Height          =   315
            Left            =   2010
            TabIndex        =   2
            Text            =   "cmb_DptDir"
            Top             =   720
            Width           =   9525
         End
         Begin VB.TextBox txt_Observ 
            Height          =   705
            Left            =   2010
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Text            =   "AteCli_frm_130.frx":000C
            Top             =   1050
            Width           =   9525
         End
         Begin VB.TextBox txt_Activi 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   9525
         End
         Begin EditLib.fpDoubleSingle ipp_IngNet 
            Height          =   315
            Left            =   2010
            TabIndex        =   0
            Top             =   60
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label lbl_General 
            Caption         =   "CIIU:"
            Height          =   285
            Index           =   39
            Left            =   90
            TabIndex        =   17
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label5 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   1050
            Width           =   1605
         End
         Begin VB.Label lbl_General 
            Caption         =   "Actividad Desarrollada:"
            Height          =   285
            Index           =   37
            Left            =   90
            TabIndex        =   9
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label lbl_General 
            Caption         =   "Ingreso Declarado (S/.):"
            Height          =   285
            Index           =   61
            Left            =   90
            TabIndex        =   8
            Top             =   60
            Width           =   1755
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
            TabIndex        =   11
            Top             =   60
            Width           =   10125
            _Version        =   65536
            _ExtentX        =   17859
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes - Actividades Económicas - Otros"
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
            Picture         =   "AteCli_frm_130.frx":0010
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   12
         Top             =   750
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
            Left            =   2010
            TabIndex        =   13
            Top             =   60
            Width           =   9525
            _Version        =   65536
            _ExtentX        =   16801
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07522154 / IKEHARA PUNK MIGUEL ANGEL (1-07521154 / IKEHARA PUNK MIGUEL ANGEL)"
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
            TabIndex        =   14
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   15
         Top             =   3090
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
            Picture         =   "AteCli_frm_130.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10200
            Picture         =   "AteCli_frm_130.frx":0624
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10890
            Picture         =   "AteCli_frm_130.frx":092E
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_CodCiu     As String
Dim l_int_FlgCmb     As Integer

Private Sub cmb_CodCiu_Change()
   l_str_CodCiu = cmb_CodCiu.Text
End Sub

Private Sub cmb_CodCiu_Click()
   If cmb_CodCiu.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Observ)
      End If
   End If
End Sub

Private Sub cmb_CodCiu_GotFocus()
   Call SendMessage(cmb_CodCiu.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
   l_str_CodCiu = cmb_CodCiu.Text
End Sub

Private Sub cmb_CodCiu_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + modgen_g_con_NUMERO + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_CodCiu, l_str_CodCiu)
      l_int_FlgCmb = True
      
      If cmb_CodCiu.ListIndex > -1 Then
         l_str_CodCiu = ""
      End If
      
      Call gs_SetFocus(txt_Observ)
   End If
End Sub

Private Sub cmb_CodCiu_LostFocus()
   Call SendMessage(cmb_CodCiu.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmd_Grabar_Click()
   If ipp_IngNet.Value = 0 Then
      MsgBox "El Ingreso Declarado no puede ser igual a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_IngNet)
      Exit Sub
   End If

   If Len(Trim(txt_Activi.Text)) = 0 Then
      MsgBox "Debe ingresar la Actividad que desarrolla el Cliente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Activi)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_Inicia_ActEco(moddat_g_int_TipCli, moddat_g_int_OrdAct)
   
   If moddat_g_int_TipCli = 1 Then
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_OrdAct = moddat_g_int_OrdAct
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_TipAct = 61
      
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Otr_IngNet = CDbl(ipp_IngNet.Text)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Otr_Activi = txt_Activi.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Otr_CodCiu = cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Otr_Observ = txt_Observ.Text
   Else
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_OrdAct = moddat_g_int_OrdAct
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_TipAct = 61
      
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Otr_IngNet = CDbl(ipp_IngNet.Text)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Otr_Activi = txt_Activi.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Otr_CodCiu = cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Otr_Observ = txt_Observ.Text
   End If
   
   moddat_g_int_FlgAct_1 = 2
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
   
   Call fs_Inicia
   Call fs_Limpia
   
   If moddat_g_int_TipCli = 1 Then
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
      If moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_TipAct = 61 Then
         ipp_IngNet.Value = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Otr_IngNet
         
         txt_Activi.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Otr_Activi
         txt_Observ.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Otr_Observ
         
         Call gs_BuscarCombo_Item(cmb_CodCiu, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Otr_CodCiu)
      End If
   Else
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli & "(" & CStr(moddat_g_int_CygTDo) & " - " & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom & ")"
   
      If moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_TipAct = 61 Then
         ipp_IngNet.Value = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Otr_IngNet
         
         txt_Activi.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Otr_Activi
         txt_Observ.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Otr_Observ
         
         Call gs_BuscarCombo_Item(cmb_CodCiu, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Otr_CodCiu)
      End If
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_CdCIIU(cmb_CodCiu)
End Sub

Private Sub fs_Limpia()
   ipp_IngNet.Value = 0
   
   txt_Activi.Text = ""
   cmb_CodCiu.ListIndex = -1
   
   txt_Observ.Text = ""
End Sub

Private Sub ipp_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Activi)
   End If
End Sub

Private Sub txt_Activi_GotFocus()
   Call gs_SelecTodo(txt_Activi)
End Sub

Private Sub txt_Activi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodCiu)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
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

