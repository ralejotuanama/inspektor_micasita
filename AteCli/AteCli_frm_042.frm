VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_SegSol_14 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   1350
   ClientTop       =   1995
   ClientWidth     =   12870
   Icon            =   "AteCli_frm_042.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5805
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12855
      _Version        =   65536
      _ExtentX        =   22675
      _ExtentY        =   10239
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
      Begin Threed.SSPanel SSPanel34 
         Height          =   2715
         Left            =   30
         TabIndex        =   5
         Top             =   2220
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   4789
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
            Height          =   855
            Left            =   1620
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   1800
            Width           =   11025
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1395
            Left            =   30
            TabIndex        =   0
            Top             =   360
            Width           =   12645
            _ExtentX        =   22304
            _ExtentY        =   2461
            _Version        =   393216
            Rows            =   12
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel35 
            Height          =   285
            Left            =   11040
            TabIndex        =   6
            Top             =   60
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Recepcionado"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel36 
            Height          =   285
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   10995
            _Version        =   65536
            _ExtentX        =   19394
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin VB.Label Label15 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   1800
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   765
         Left            =   30
         TabIndex        =   9
         Top             =   4980
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            Left            =   12000
            Picture         =   "AteCli_frm_042.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_RecDoc 
            Height          =   675
            Left            =   11310
            Picture         =   "AteCli_frm_042.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            Height          =   615
            Left            =   630
            TabIndex        =   11
            Top             =   30
            Width           =   4365
            _Version        =   65536
            _ExtentX        =   7699
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "Recepción de Documentos"
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
            Left            =   5190
            TabIndex        =   12
            Top             =   120
            Width           =   7485
            _Version        =   65536
            _ExtentX        =   13203
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
            Left            =   60
            Picture         =   "AteCli_frm_042.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   13
         Top             =   750
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            TabIndex        =   14
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "001-001-04-0001"
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
         Begin Threed.SSPanel pnl_EjeVta 
            Height          =   315
            Left            =   1620
            TabIndex        =   15
            Top             =   1050
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            TabIndex        =   16
            Top             =   720
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "BIEN TERMINADO"
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
            TabIndex        =   17
            Top             =   390
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO - MIVIVIENDA"
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
            Left            =   8820
            TabIndex        =   18
            Top             =   60
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "DOLARES AMERICANOS"
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
         Begin VB.Label Label24 
            Caption         =   "Moneda Prést.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   23
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   22
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_RecDoc_Click()
   Dim r_int_FlgDoc     As Integer
   Dim r_int_Contad     As Integer
   Dim r_str_Cadena     As String
   Dim r_str_Selecc     As String
   
   'Validando Documentos a Recibir
   r_int_FlgDoc = 1
   
   r_str_Selecc = ""
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 1
      r_str_Selecc = Trim(grd_Listad.Text)
      
      grd_Listad.Col = 6
      
      'Si es de Obligatoria Selección
      If Len(Trim(r_str_Selecc)) = 0 And CInt(grd_Listad.Text) = 1 Then
         r_int_FlgDoc = 2
         Exit For
      End If
   Next r_int_Contad
   Call gs_UbiIniGrid(grd_Listad)
   
   If r_int_FlgDoc = 2 Then
      MsgBox "Debe seleccionar los Documentos que han sido recibidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If ff_Grabar() Then
      moddat_g_int_FlgAct = 2
      
      r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      r_str_Cadena = r_str_Cadena & Chr(13)
   
      modgen_g_str_Mail_Asunto = "RECEPCION DE DOCUMENTOS (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      modgen_g_str_Mail_Mensaj = r_str_Cadena
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & txt_Observ.Text
      
      frm_EnvMai_01.Show 1
      
      Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   txt_Observ.Text = ""
   
   Call fs_Inicia
   Call fs_Carga_DatGen
   Call fs_Carga_Docume
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Carga_DatGen()
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod
   pnl_EjeVta.Caption = moddat_g_str_EjeVta
   pnl_Moneda.Caption = moddat_g_str_Moneda
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 10985
   grd_Listad.ColWidth(1) = 1300
   grd_Listad.ColWidth(2) = 0
   grd_Listad.ColWidth(3) = 0
   grd_Listad.ColWidth(4) = 0
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColWidth(6) = 0

   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
End Sub

Private Sub fs_Carga_Docume()
   Call gs_LimpiaGrid(grd_Listad)
   
   'Documentos Legales
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '302' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARPRD_DESCRI)
         grd_Listad.Col = 1:     grd_Listad.Text = ""
         grd_Listad.Col = 2:     grd_Listad.Text = "1"
         grd_Listad.Col = 3:     grd_Listad.Text = "302"
         grd_Listad.Col = 4:     grd_Listad.Text = "0"
         grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARPRD_CODITE
         grd_Listad.Col = 6:     grd_Listad.Text = Left(g_rst_Genera!PARPRD_DESCRI, 1)
         
         g_rst_Genera.MoveNext
      Loop
      
      grd_Listad.Redraw = True
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   'Documentos del Vendedor
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '303' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARPRD_DESCRI)
         grd_Listad.Col = 1:     grd_Listad.Text = ""
         grd_Listad.Col = 2:     grd_Listad.Text = "1"
         grd_Listad.Col = 3:     grd_Listad.Text = "303"
         grd_Listad.Col = 4:     grd_Listad.Text = "0"
         grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARPRD_CODITE
         grd_Listad.Col = 6:     grd_Listad.Text = Left(g_rst_Genera!PARPRD_DESCRI, 1)
         
         g_rst_Genera.MoveNext
      Loop
      
      grd_Listad.Redraw = True
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   'Documentos por Modalidad
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '304' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(PARPRD_CODITE,1,1) = '" & Format(CInt(moddat_g_str_CodMod), "0") & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARPRD_DESCRI)
         grd_Listad.Col = 1:     grd_Listad.Text = ""
         grd_Listad.Col = 2:     grd_Listad.Text = "1"
         grd_Listad.Col = 3:     grd_Listad.Text = "304"
         grd_Listad.Col = 4:     grd_Listad.Text = "0"
         grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARPRD_CODITE
         grd_Listad.Col = 6:     grd_Listad.Text = Left(g_rst_Genera!PARPRD_DESCRI, 1)
         
         g_rst_Genera.MoveNext
      Loop
      
      grd_Listad.Redraw = True
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   If grd_Listad.Rows > 0 Then
      Call gs_UbiIniGrid(grd_Listad)
   Else
      MsgBox "No se encontró la Lista de Documentos a Recibir.", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 1
      
      If grd_Listad.Text = "X" Then
         grd_Listad.Text = ""
      Else
         grd_Listad.Text = "X"
      End If
      
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub grd_Listad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then
      Call grd_Listad_DblClick
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_RecDoc)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Function ff_Grabar() As Integer
   Dim r_int_Contad     As Integer
   Dim r_str_FlgDoc     As String
   Dim r_int_TipDoc     As Integer
   Dim r_str_CodGrp     As String
   Dim r_int_CodAct     As Integer
   Dim r_str_CodIte     As String
   
   Call moddat_gs_FecSis
   
   ff_Grabar = False
   
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaTas, 23, 0, txt_Observ.Text, 0, 0) Then
      Exit Function
   End If
   
   'Grabando Documentos Recibidos
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 1
      r_str_FlgDoc = grd_Listad.Text
      
      grd_Listad.Col = 2
      r_int_TipDoc = CInt(grd_Listad.Text)
      
      grd_Listad.Col = 3
      r_str_CodGrp = grd_Listad.Text
      
      grd_Listad.Col = 4
      r_int_CodAct = CInt(grd_Listad.Text)
      
      grd_Listad.Col = 5
      r_str_CodIte = grd_Listad.Text
      
      If r_str_FlgDoc = "X" Then
         If Not moddat_gf_Inserta_SolDoc(moddat_g_str_NumSol, r_int_TipDoc, moddat_g_str_CodPrd, moddat_g_str_CodSub, r_int_CodAct, r_str_CodGrp, r_str_CodIte, Format(CDate(moddat_g_str_FecSis), "yyyymmdd")) Then
            Exit Function
         End If
      End If
   Next r_int_Contad
   
   ff_Grabar = True
End Function

