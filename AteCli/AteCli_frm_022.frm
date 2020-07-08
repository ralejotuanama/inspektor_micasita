VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_RecAdm_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   1830
   ClientTop       =   1350
   ClientWidth     =   11640
   Icon            =   "AteCli_frm_022.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7485
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   13203
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
         TabIndex        =   22
         Top             =   6660
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_022.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   570
            Top             =   0
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
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1545
         Left            =   30
         TabIndex        =   13
         Top             =   5070
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   2725
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
         Begin VB.ComboBox cmb_MotRec 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   60
            Width           =   9855
         End
         Begin VB.TextBox txt_Observ 
            Height          =   1125
            Left            =   1620
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Text            =   "AteCli_frm_022.frx":044E
            Top             =   390
            Width           =   9855
         End
         Begin VB.Label Label8 
            Caption         =   "Motivo Rechazo:"
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label Label23 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   390
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   11
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
            TabIndex        =   12
            Top             =   60
            Width           =   4905
            _Version        =   65536
            _ExtentX        =   8652
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Rechazo Administrativo"
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
            Picture         =   "AteCli_frm_022.frx":0452
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   795
         Left            =   30
         TabIndex        =   16
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1402
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   9450
            Picture         =   "AteCli_frm_022.frx":0D1C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_022.frx":1026
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_022.frx":1330
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   6210
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   2775
         End
         Begin MSMask.MaskEdBox msk_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   3
            Top             =   390
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "###-###-##-####"
            PromptChar      =   " "
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   4830
            TabIndex        =   19
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   4830
            TabIndex        =   18
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Solicitud:"
            Height          =   285
            Left            =   90
            TabIndex        =   17
            Top             =   390
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   23
         Top             =   1590
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
            Left            =   1620
            TabIndex        =   24
            Top             =   60
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   25
            Top             =   390
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
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
            Left            =   8040
            TabIndex        =   26
            Top             =   30
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
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   29
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label10 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   28
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Ingreso:"
            Height          =   315
            Left            =   6840
            TabIndex        =   27
            Top             =   60
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2625
         Left            =   30
         TabIndex        =   30
         Top             =   2400
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   4630
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   2505
            Left            =   60
            TabIndex        =   31
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   4419
            _Version        =   393216
            Style           =   1
            Tabs            =   6
            TabsPerRow      =   6
            TabHeight       =   520
            TabCaption(0)   =   "Datos del Cliente"
            TabPicture(0)   =   "AteCli_frm_022.frx":1772
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos del Cónyuge"
            TabPicture(1)   =   "AteCli_frm_022.frx":178E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Patrimonio"
            TabPicture(2)   =   "AteCli_frm_022.frx":17AA
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(4)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Referencias Personales"
            TabPicture(3)   =   "AteCli_frm_022.frx":17C6
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos del Inmueble"
            TabPicture(4)   =   "AteCli_frm_022.frx":17E2
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(2)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Datos del Crédito"
            TabPicture(5)   =   "AteCli_frm_022.frx":17FE
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(6)"
            Tab(5).Control(1)=   "grd_Listad(5)"
            Tab(5).ControlCount=   2
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2055
               Index           =   0
               Left            =   60
               TabIndex        =   32
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3625
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2055
               Index           =   2
               Left            =   -74940
               TabIndex        =   33
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3625
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2055
               Index           =   3
               Left            =   -74940
               TabIndex        =   34
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3625
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2055
               Index           =   4
               Left            =   -74940
               TabIndex        =   35
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3625
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1125
               Index           =   5
               Left            =   -74940
               TabIndex        =   36
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   1984
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   915
               Index           =   6
               Left            =   -74940
               TabIndex        =   37
               Top             =   1530
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   1614
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2055
               Index           =   1
               Left            =   -74940
               TabIndex        =   38
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3625
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
         End
      End
   End
End
Attribute VB_Name = "frm_RecAdm_01"
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

Private Sub cmd_Buscar_Click()
   If cmb_TipBus.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Búsqueda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipBus)
      Exit Sub
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      If cmb_TipDoc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipDoc)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumDoc.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
         txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
      End If
      
      moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_str_TipDoc = cmb_TipDoc.Text
      moddat_g_str_NumDoc = txt_NumDoc.Text
   Else
      If Len(Trim(msk_NumSol.Text)) < 12 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      moddat_g_str_NumSol = msk_NumSol.Text
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = '" & moddat_g_str_NumDoc & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   Else
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe Solicitud en Trámite para la Selección de Búsqueda. ", vbExclamation, modgen_g_str_NomPlt
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call cmd_Limpia_Click
      Exit Sub
   End If

   Call fs_Buscar_DatGen

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_FecIng.Caption = moddat_g_str_FecIng
   
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)
   Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
   Call fs_DatPat
   Call fs_DatRef
   Call fs_DatInm
   Call fs_DatCre
   Call fs_SolDoc
   

   Call fs_Activa(False)
   Call gs_SetFocus(cmb_MotRec)
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_MotRec.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Motivo del Rechazo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MotRec)
      Exit Sub
   End If
   
   If Len(Trim(txt_Observ.Text)) = 0 Then
      MsgBox "Debe ingresar las Observaciones del Rechazo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Observ)
      Exit Sub
   End If
   
   
   If MsgBox("Al Rechazar Administrativamente una Solicitud cancela todo el proceso de Tramitación. ¿Está seguro de Rechazar la Solicitud?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   'Generando Ocurrencia en Detalle de Seguimientos
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 91, 14, 0, "", 0, cmb_MotRec.ItemData(cmb_MotRec.ListIndex)) Then
      Exit Sub
   End If
      
   'Creando Registro en Cabecera de Seguimiento
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, 91) Then
      Exit Sub
   End If
      
   'Modificando Registro en Cabecera de Seguimiento
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 91, 0, 2, 1) Then
      Exit Sub
   End If
      
   'Registrando Información del Rechazo
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_TRA_RECADM ("
   
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_MotRec.ItemData(cmb_MotRec.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_TRA_RECADM. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop

   modgen_g_str_Mail_Asunto = "RECHAZO ADMINISTRATIVO (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_gf_Consulta_ParDes("003", moddat_g_int_MotRec)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & txt_Observ.Text
   
   Call fs_Envia_CorEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)

   MsgBox "La solicitud fue rechazada con éxito.", vbInformation, modgen_g_str_NomPlt
   Call cmd_Limpia_Click
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmb_TipBus_Click()
   If cmb_TipBus.ListIndex > -1 Then
      If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
         cmb_TipDoc.Enabled = True
         txt_NumDoc.Enabled = True
         msk_NumSol.Enabled = False
         
         msk_NumSol.Mask = ""
         msk_NumSol.Text = ""
         msk_NumSol.Mask = "###-###-##-####"
         
         Call gs_SetFocus(cmb_TipDoc)
      Else
         cmb_TipDoc.Enabled = False
         txt_NumDoc.Enabled = False
         msk_NumSol.Enabled = True
         
         cmb_TipDoc.ListIndex = -1
         txt_NumDoc.Text = ""
         
         Call gs_SetFocus(msk_NumSol)
      End If
   Else
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      
      msk_NumSol.Enabled = False
   
      cmb_TipDoc.ListIndex = -1
      txt_NumDoc.Text = ""
      msk_NumSol.Mask = ""
      msk_NumSol.Text = ""
      msk_NumSol.Mask = "###-###-##-####"
   End If
End Sub

Private Sub cmb_TipBus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipBus_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer
   
   Call modsis_gs_Carga_TipBus(cmb_TipBus)
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   
   Call moddat_gs_Carga_MotRec(cmb_MotRec, 91)

   'Inicializando Grid de Cliente y de Cónyuge
   For r_int_Contad = 0 To 5
      grd_Listad(r_int_Contad).ColWidth(0) = 3000
      grd_Listad(r_int_Contad).ColWidth(1) = 7940
   
      grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
   Next r_int_Contad
   
   grd_Listad(6).ColWidth(0) = 10940
   grd_Listad(6).ColAlignment(0) = flexAlignLeftCenter
End Sub

Private Sub msk_NumSol_GotFocus()
   Call gs_SelecTodo(msk_NumSol)
End Sub

Private Sub msk_NumSol_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub fs_Limpia()
   Dim r_int_Contad     As Integer
   
   Call fs_Activa(True)
   
   cmb_TipBus.ListIndex = -1
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   msk_NumSol.Enabled = False

   msk_NumSol.Mask = ""
   msk_NumSol.Text = ""
   msk_NumSol.Mask = "###-###-##-####"
   
   txt_NumDoc.Text = ""
   
   pnl_Produc.Caption = ""
   pnl_NumSol.Caption = ""
   pnl_FecIng.Caption = ""
   
   cmb_MotRec.ListIndex = -1
   txt_Observ.Text = ""

   For r_int_Contad = 0 To 5
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   Dim r_int_Contad     As Integer
   
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumSol.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   cmd_Grabar.Enabled = Not p_Habilita
   cmb_MotRec.Enabled = Not p_Habilita
   txt_Observ.Enabled = Not p_Habilita
   
   For r_int_Contad = 0 To 6
      grd_Listad(r_int_Contad).Enabled = Not p_Habilita
   Next r_int_Contad
End Sub

Private Sub fs_Buscar_DatGen()
   g_rst_Princi.MoveFirst

   'Datos de Búsqueda
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO)
   moddat_g_str_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(CStr(moddat_g_int_TipDoc), Trim(moddat_g_str_NumDoc))
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(moddat_g_str_CodPrd, moddat_g_str_CodMod)

   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG)
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(moddat_g_str_CodEjeSeg)
   
   'Consejero
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP)
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(moddat_g_str_CodConHip)
   
   'Instancia Actual
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS
   moddat_g_str_InsAct = moddat_gf_Consulta_ParDes("002", Trim(moddat_g_int_InsAct))
   
   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))

   'Fecha de Ingreso
   moddat_g_str_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
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

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygTDo = ""

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(p_Indice).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      If p_Indice = 1 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Documento de Identidad"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DatGen_NumDoc & "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Apellidos y Nombres"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DatGen_Nombre)
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Documento Adicional de Identidad"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DatGen_FLGDOA)) & IIf(g_rst_Princi!DatGen_FLGDOA = 1, " ( " & moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TIPDOA)) & " - " & Trim(g_rst_Princi!DatGen_NUMDOA) & ")", "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Sexo"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("207", CStr(g_rst_Princi!DatGen_CodSex))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Fecha de Nacimiento"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Nacionalidad"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Lugar de Nacimiento"
   
      If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      Else
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = "<< NO REGISTRADO >>"
      End If
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Estado Civil"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DatGen_RegCyg), "")
         
         If g_rst_Princi!DATGEN_ESTCIV = 2 Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
            moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
            moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
         End If
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Nivel de Estudios"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Profesión"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("501", CStr(g_rst_Princi!DatGen_Profes))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Celular"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Nro. Dependientes Económicos"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = CStr(g_rst_Princi!DatGen_DepEco)
      
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Edades"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = IIf(g_rst_Princi!DatGen_EDAD01 > 0, CStr(g_rst_Princi!DatGen_EDAD01), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD02 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD02), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD03 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD03), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD04 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD04), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD05 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD05), "")
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "E-mail"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_DirEle & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Autorización Envío"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_AUTENV))
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Domicilio"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Princi!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_Refere & "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Teléfono Domicilio"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_Telefo & "")
      End If
      
      grd_Listad(p_Indice).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(p_Indice))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Actividad Económica Principal
   Call fs_ActEco(p_TipDoc, p_NumDoc, 1, p_Indice)
   Call fs_ActEco(p_TipDoc, p_NumDoc, 2, p_Indice)
End Sub

Private Sub fs_ActEco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_OrdAct As Integer, ByVal p_Indice As Integer)
   Dim r_var_ColTxt
   
   If p_OrdAct = 1 Then
      r_var_ColTxt = modgen_g_con_ColAzu
   Else
      r_var_ColTxt = modgen_g_con_ColRoj
   End If

   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(p_Indice).Redraw = False
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = r_var_ColTxt
      grd_Listad(p_Indice).Text = "Ocupación " & Left(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct)), 1) & Mid(LCase(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct))), 2)
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = r_var_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("008", g_rst_Princi!ActEco_CodAct)
      
      Select Case g_rst_Princi!ActEco_CodAct
         Case 11: Call fs_ActEco_Dep(p_Indice, r_var_ColTxt)
         Case 21: Call fs_ActEco_Ind(p_Indice, r_var_ColTxt)
         Case 31: Call fs_ActEco_Com(p_Indice, r_var_ColTxt)
         Case 41: Call fs_ActEco_Acc(p_Indice, r_var_ColTxt)
         Case 51: Call fs_ActEco_Ren(p_Indice, r_var_ColTxt)
      End Select
      
      grd_Listad(p_Indice).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(p_Indice))
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_ActEco_Dep(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad Empleador"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Situación como Trabajador"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("235", g_rst_Princi!ActEco_Dep_SitTra)

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_RazSoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Comercial"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomCom & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Dep_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Dep_CodCiu))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TeleRH & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_AnexRH & "")) > 0, " ANEXO: " & Trim(g_rst_Princi!ActEco_Dep_AnexRH & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
   Else
      g_rst_Genera.MoveFirst

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Genera!DATGEN_COCIIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_COCIIU))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELERH & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELERH & "")) > 0, " ANEXO: " & Trim(g_rst_Genera!DATGEN_ANEXRH & ""), "")

      If g_rst_Princi!ActEco_Dep_TipOfi = 1 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Dirección"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Genera!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_Refere & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Fax"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_NUMFAX & "")
      Else
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Dirección"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                     " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                     IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Fax"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
      End If
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Dep_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Frecuencia de Haberes"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("210", CStr(g_rst_Princi!ActEco_Dep_FreHab))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Ingreso"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Cargo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Dep_CodCar = "999999", Trim(g_rst_Princi!ActEco_Dep_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Dep_CodCar))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Area"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomAre & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Anexo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono Directo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Celular"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Celula & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "E-mail"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_DirEle & "")

   If g_rst_Princi!ActEco_Dep_TraAnt = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Documento Identidad Empleador Anterior"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc_Ant) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc_Ant & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc_Ant) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc_Ant & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social (Empleador Anterior)"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_RazSoc_Ant & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Nombre Comercial (Empleador Anterior)"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomCom_Ant & "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s) (Empleador Anterior)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1_Ant & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      End If

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Ingreso (Empleador Anterior)"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng_Ant))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Cese (Empleador Anterior)"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecCes_Ant))
   End If
End Sub

Private Sub fs_ActEco_Ind(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Dirección"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Ind_TipVia)) & _
                               " " & Trim(g_rst_Princi!ActEco_Ind_NomVia) & " " & Trim(g_rst_Princi!ActEco_Ind_NumVia) & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Ind_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Ind_IntDpt) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Ind_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Ind_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Ind_NomZon), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Referencia"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Refere & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Indartamento / Provincia / Distrito"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ind_UbiGeo))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2 & ""), "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fax"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Ind_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Ind_CodCiu))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ind_IngNet, 15, 2)
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Actividades"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Contrato de Locación de Servicios"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
   
   If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Documento Identidad Empleador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Ind_NumDoc_Emp & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_RazSoc_Emp & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Nombre Comercial"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_NomCom_Emp & "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Telef1_Emp & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      End If

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Ingreso"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Cargo"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Ind_CodCar = "999999", Trim(g_rst_Princi!ActEco_Ind_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Ind_CodCar))
   End If
End Sub

Private Sub fs_ActEco_Com(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Com_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Com_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Razón Social"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Nombre Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Dirección"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Com_TipVia)) & _
                               " " & Trim(g_rst_Princi!ActEco_Com_NomVia) & " " & Trim(g_rst_Princi!ActEco_Com_NumVia) & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Com_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Com_IntDpt) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Com_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Com_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Com_NomZon), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Referencia"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_Refere & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Com_UbiGeo))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Com_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Com_Telef2 & ""), "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fax"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NumFax & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Com_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Com_CodCiu))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Giro Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_GirCom(g_rst_Princi!ActEco_Com_GirCom)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ventas Mensuales"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_VtaMen, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Operaciones"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Cargo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Com_CodCar = "999999", Trim(g_rst_Princi!ActEco_Com_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Com_CodCar))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Régimen Tributario"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("215", CStr(g_rst_Princi!ActEco_Com_RegTri))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participación"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_PorPar, 7, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tipo de Local Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("208", CStr(g_rst_Princi!ActEco_Com_TipLoc))
   
   If g_rst_Princi!ActEco_Com_TipLoc = 2 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Alquiler Mensual"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_AlqMen, 15, 2)
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Arrendador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NomArr & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono Arrendador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_TelArr & "")
   End If
End Sub

Private Sub fs_ActEco_Acc(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Acc_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Acc_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Acc_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_RazSoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Comercial"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_NomCom & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Acc_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Acc_CodCiu))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Acc_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Acc_NomVia) & " " & Trim(g_rst_Princi!ActEco_Acc_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Acc_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Acc_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Acc_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Acc_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Acc_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Acc_UbiGeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Acc_UbiGeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Acc_UbiGeo))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Acc_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_NumFax & "")
   Else
      g_rst_Genera.MoveFirst

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Genera!DATGEN_COCIIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_COCIIU))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                  " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                  IIf(Len(Trim(g_rst_Genera!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Genera!DatGen_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_NUMFAX & "")
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Antigüedad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participación"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_PorPar, 7, 2)
End Sub

Private Sub fs_ActEco_Ren(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Dirección de Propiedad 01"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Nombre de Arrendatario"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Alquiler"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele21 & ""), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Alquiler Mensual"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe1, 15, 2)
   
   If g_rst_Princi!ActEco_Ren_SegPro = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección de Propiedad 02"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre de Arrendatario"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Inicio de Alquiler"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele22 & ""), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Alquiler Mensual"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe2, 15, 2)
   End If
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub fs_DatInm()
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(2).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Tipo de Inmueble"
         
      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("217", CStr(g_rst_Princi!SOLINM_TIPINM))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Dirección"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON), "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Referencia"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Departamento / Provincia / Distrito"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Proyecto miCasita"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
      
      If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(2).Text = "Nombre Proyecto"

         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(2).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
      Else
         If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = "Nombre Proyecto"

            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
         End If
      End If
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = "Propietario / Promotor"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("218", g_rst_Princi!SOLINM_FLGPRO)
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = "Docum. Identidad Propietario/Promotor"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = "Nombre o Razón Social"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = "Dirección"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = "Referencia"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = "Departamento / Provincia / Distrito"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = "Teléfono"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
      
      If g_rst_Princi!SOLINM_FLGCON = 1 Then
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = "Docum. Identidad Constructor"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_CON)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = "Nombre o Razón Social"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = "Dirección"
         
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_CON)) & _
                           " " & Trim(g_rst_Princi!SOLINM_NOMVIA_CON) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_CON) & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_CON)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_CON) & ")", "") & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_CON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_CON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_CON), "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = "Referencia"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2) & "0000") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 4) & "00") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_CON))
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = "Teléfono"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
      End If
      
      grd_Listad(2).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(2))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatRef()
   Dim r_var_ColTxt

   r_var_ColTxt = modgen_g_con_ColNeg

   Call gs_LimpiaGrid(grd_Listad(3))

   g_str_Parame = "SELECT * FROM CRE_SOLREF WHERE "
   g_str_Parame = g_str_Parame & "SOLREF_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(3).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Tipo de Referencia"
            
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("010", CStr(g_rst_Princi!SOLREF_TIPREF))
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Tipo de Parentesco"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         
         If g_rst_Princi!SOLREF_TIPREF = 1 Then
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("212", CStr(g_rst_Princi!SOLREF_TIPPAR))
         Else
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("213", CStr(g_rst_Princi!SOLREF_TIPPAR))
         End If
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Apellidos y Nombres"
   
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_APEPAT & "") & " " & Trim(g_rst_Princi!SOLREF_APEMAT & "") & " " & Trim(g_rst_Princi!SOLREF_NOMBRE & "")
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Teléfono"

         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_TELEFO & "")
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Celular"
   
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_CELULA & "")
   
         g_rst_Princi.MoveNext
         
         If r_var_ColTxt = modgen_g_con_ColNeg Then
            r_var_ColTxt = modgen_g_con_ColAzu
         Else
            r_var_ColTxt = modgen_g_con_ColNeg
         End If
      Loop
      
      grd_Listad(3).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(3))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatPat()
   Dim r_int_Contad     As Integer
   
   Call gs_LimpiaGrid(grd_Listad(4))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   grd_Listad(4).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   If g_rst_Princi!SOLMAE_REGIMB = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "INMUEBLES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLINB WHERE "
      g_str_Parame = g_str_Parame & "SOLINB_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLINB_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Tipo Inmueble (" & Format(r_int_Contad, "00") & ")"
   
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("216", CStr(g_rst_Genera!SOLINB_TIPINM))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Fecha de Adquisición (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoFecha(CStr(g_rst_Genera!SOLINB_FECADQ))
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Importe Valorizado (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLINB_IMPVAL, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Dirección (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLINB_DIRECC & "")
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "INMUEBLES"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   
   If g_rst_Princi!SOLMAE_REGTAR = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "TARJETAS DE CREDITO"
      
      g_str_Parame = "SELECT * FROM CRE_SOLTRJ WHERE "
      g_str_Parame = g_str_Parame & "SOLTRJ_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLTRJ_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Institución Financiera (" & Format(r_int_Contad, "00") & ")"
   
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLTRJ_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Tipo de Tarjeta (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("506", g_rst_Genera!SOLTRJ_TIPTRJ)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Número de Tarjeta (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLTRJ_NUMTRJ & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLTRJ_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Saldo Actual (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_SALACT, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Línea Crédito (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_LIMCRD, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Pago Mínimo (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_PAGMIN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "TARJETAS DE CREDITO"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGDEU = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "DEUDAS"
      
      g_str_Parame = "SELECT * FROM CRE_SOLDEU WHERE "
      g_str_Parame = g_str_Parame & "SOLDEU_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLDEU_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Institución Financiera (" & Format(r_int_Contad, "00") & ")"
   
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLDEU_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Número de Operación (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLDEU_NUMOPE & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLDEU_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Monto del Préstamo (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_MTOOTO, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Saldo por Pagar (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_SALPAG, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Cuota Mensual (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_CUOMEN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Meses x Pagar (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = CStr(g_rst_Genera!SOLDEU_PLAMEN)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "DEUDAS"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGGAS = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "GASTOS MENSUALES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLEYM WHERE "
      g_str_Parame = g_str_Parame & "SOLEYM_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLEYM_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("220", g_rst_Genera!SOLEYM_CODEYM)
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLEYM_IMPORT, 12, 2)
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "GASTOS MENSUALES"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "NO REGISTRA"
   End If
   
   grd_Listad(4).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(4))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   Call gs_LimpiaGrid(grd_Listad(5))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   grd_Listad(5).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Sub-Producto"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Moneda del Préstamo"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Fecha de Solicitud"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Tasa de Interés"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"
   
   If g_rst_Princi!SOLMAE_INMIDE = 1 Then
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Modalidad"
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!SOLMAE_CODMOD)), "000")) Then
         grd_Listad(5).Col = 1
         grd_Listad(5).Text = moddat_g_arr_Genera(1).Genera_Nombre
      End If
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Moneda de Compra-Venta"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_COMVTA_MON))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Valor de Compra Venta (US$)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Aporte Propio (US$)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Préstamo (US$)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Valor de Compra Venta (S/.)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Aporte Propio (S/.)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Préstamo (S/.)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tipo de Cambio Referencial"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL / g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Plazo (Años)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_PLAANO)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Período de Gracia (Meses)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_PERGRA)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Cuotas Extraordinarias"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_CUOEXT))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Compañía de Seguros"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tipo de Seguro Desgravamen"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Día de Pago"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   End If
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Observaciones"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Consejero Hipotecario"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Ejecutivo de Seguimiento"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_EJESEG)
   
   
   grd_Listad(5).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(5))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_SolDoc()
   Call gs_LimpiaGrid(grd_Listad(6))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLDOC WHERE "
   g_str_Parame = g_str_Parame & "SOLDOC_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "(SOLDOC_TIPDOC = 1 OR SOLDOC_TIPDOC = 2)"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(6).Redraw = False
   Do While Not g_rst_Princi.EOF
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
   
      grd_Listad(6).Col = 0
      
      If g_rst_Princi!SOLDOC_TIPDOC = 1 Then
         'Buscar en Parámetros por Producto
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(6).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      Else
         'Buscar en Parámetros por Actividad Económica
         If moddat_gf_Consulta_ParAct(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, CStr(g_rst_Princi!SOLDOC_CODACT), g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(6).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad(6).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(6))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Envia_CorEle(ByVal p_Asunto As String, ByVal p_Mensaje As String)
   Dim r_str_Cadena     As String
   
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Jefe de Seguimiento
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(130)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Ventas
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(120)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director Comercial
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(100)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Area según Instancia Actual
   Select Case moddat_g_int_InsAct
      Case 21, 71
         r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(210)    'Jefe de Créditos
      Case 41, 42, 61, 62, 72, 81
         r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(220)    'Jefe de Operaciones
      Case 51
         r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(230)    'Jefe de Legal
   End Select
   
   'Director de Producción
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(200)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
End Sub


