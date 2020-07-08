VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_LisRec_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   2550
   ClientTop       =   5985
   ClientWidth     =   14790
   Icon            =   "AteCli_frm_059.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      _Version        =   65536
      _ExtentX        =   26061
      _ExtentY        =   7329
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
         Height          =   675
         Left            =   30
         TabIndex        =   13
         Top             =   750
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14100
            Picture         =   "AteCli_frm_059.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2145
         Left            =   30
         TabIndex        =   1
         Top             =   1950
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
         _ExtentY        =   3784
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1785
            Left            =   30
            TabIndex        =   2
            Top             =   330
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   3149
            _Version        =   393216
            Rows            =   21
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   3300
            TabIndex        =   4
            Top             =   60
            Width           =   2955
            _Version        =   65536
            _ExtentX        =   5212
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Rechazo"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   1620
            TabIndex        =   5
            Top             =   60
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Rechazo"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   12570
            TabIndex        =   6
            Top             =   60
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Cliente"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   6240
            TabIndex        =   7
            Top             =   60
            Width           =   6345
            _Version        =   65536
            _ExtentX        =   11192
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Motivo de Rechazo"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
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
            Left            =   660
            TabIndex        =   12
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Lista de Solicitudes Rechazadas"
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
            Left            =   90
            Picture         =   "AteCli_frm_059.frx":044E
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   435
         Left            =   30
         TabIndex        =   9
         Top             =   1470
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1860
            TabIndex        =   10
            Top             =   60
            Width           =   12795
            _Version        =   65536
            _ExtentX        =   22569
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
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frm_LisRec_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_int_Contad     As Integer
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   pnl_Client.Caption = CStr(moddat_g_int_CygTDo) & "-" & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom

   grd_Listad.ColWidth(0) = 1560
   grd_Listad.ColWidth(1) = 1680
   grd_Listad.ColWidth(2) = 2940
   grd_Listad.ColWidth(3) = 6360
   grd_Listad.ColWidth(4) = 1800
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   
   Call gs_LimpiaGrid(grd_Listad)
   
   For r_int_Contad = 1 To UBound(modatecli_g_arr_CygRec)
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Left(modatecli_g_arr_CygRec(r_int_Contad).LisRec_NumSol, 3) & "-" & Mid(modatecli_g_arr_CygRec(r_int_Contad).LisRec_NumSol, 4, 3) & "-" & Mid(modatecli_g_arr_CygRec(r_int_Contad).LisRec_NumSol, 7, 2) & "-" & Right(modatecli_g_arr_CygRec(r_int_Contad).LisRec_NumSol, 4)
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(CStr(modatecli_g_arr_CygRec(r_int_Contad).LisRec_FecRec))
      
      grd_Listad.Col = 2
      grd_Listad.Text = moddat_gf_Consulta_ParDes("021", modatecli_g_arr_CygRec(r_int_Contad).LisRec_TipRec)
      
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Consulta_ParDes("003", modatecli_g_arr_CygRec(r_int_Contad).LisRec_MotRec)
      
      grd_Listad.Col = 4
      grd_Listad.Text = moddat_gf_Consulta_ParDes("014", modatecli_g_arr_CygRec(r_int_Contad).LisRec_TipCli)
   Next r_int_Contad

   Call gs_UbiIniGrid(grd_Listad)
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub


