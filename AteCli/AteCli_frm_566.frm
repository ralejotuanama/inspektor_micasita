VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_PryNVi_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9390
   Icon            =   "AteCli_frm_566.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   7635
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9375
      _Version        =   65536
      _ExtentX        =   16536
      _ExtentY        =   13467
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   615
         Left            =   60
         TabIndex        =   10
         Top             =   765
         Width           =   9255
         _Version        =   65536
         _ExtentX        =   16325
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
         Begin Threed.SSPanel pnl_Proyecto 
            Height          =   375
            Left            =   1590
            TabIndex        =   0
            Top             =   120
            Width           =   7575
            _Version        =   65536
            _ExtentX        =   13361
            _ExtentY        =   661
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Proyecto :"
            Height          =   195
            Left            =   105
            TabIndex        =   11
            Top             =   165
            Width           =   1320
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3855
         Left            =   60
         TabIndex        =   12
         Top             =   1410
         Width           =   9255
         _Version        =   65536
         _ExtentX        =   16325
         _ExtentY        =   6800
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
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   7320
            Picture         =   "AteCli_frm_566.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Adicionar Registro de Limite de RRPP"
            Top             =   3120
            Width           =   585
         End
         Begin VB.CommandButton cmd_Modifica 
            Height          =   585
            Left            =   7920
            Picture         =   "AteCli_frm_566.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Modificar Registro de Limite de RRPP"
            Top             =   3120
            Width           =   585
         End
         Begin VB.CommandButton cmd_Elimina 
            Height          =   585
            Left            =   8520
            Picture         =   "AteCli_frm_566.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Eliminar Registro de Limite de RRPP"
            Top             =   3120
            Width           =   585
         End
         Begin MSFlexGridLib.MSFlexGrid grdLstRRPP 
            Height          =   2865
            Left            =   120
            TabIndex        =   1
            Top             =   120
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   5054
            _Version        =   393216
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   1575
         Left            =   60
         TabIndex        =   13
         Top             =   5300
         Width           =   9255
         _Version        =   65536
         _ExtentX        =   16325
         _ExtentY        =   2778
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
         Begin VB.TextBox txt_ComenRP 
            Height          =   975
            Left            =   1200
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   480
            Width           =   7935
         End
         Begin EditLib.fpDateTime ipp_FecInfInm 
            Height          =   315
            Left            =   1200
            TabIndex        =   5
            Top             =   120
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   -1  'True
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
            Text            =   ""
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Comentarios :"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Limite :"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   990
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   690
         Left            =   60
         TabIndex        =   16
         Top             =   6900
         Width           =   9255
         _Version        =   65536
         _ExtentX        =   16325
         _ExtentY        =   1217
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
            Height          =   585
            Left            =   8040
            Picture         =   "AteCli_frm_566.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Grabar datos"
            Top             =   60
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   8640
            Picture         =   "AteCli_frm_566.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   60
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   9255
         _Version        =   65536
         _ExtentX        =   16325
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
         Begin Threed.SSPanel pnl_TituloForm 
            Height          =   495
            Left            =   630
            TabIndex        =   18
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Gestión de Proyectos No Vinculados"
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
            Picture         =   "AteCli_frm_566.frx":11AE
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_PryNVi_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_Mant As String

Private Sub cmd_Agrega_Click()
   l_str_Mant = "N"

   cmd_Agrega.Enabled = False
   cmd_Modifica.Enabled = False
   cmd_Elimina.Enabled = False
   cmd_Grabar.Enabled = True
   
   ipp_FecInfInm.Enabled = True
   txt_ComenRP.Enabled = True
   
   Call gs_SetFocus(ipp_FecInfInm)
End Sub

Private Sub cmd_Elimina_Click()
   If MsgBox("¿Está seguro de eliminar la Fecha Limite asignada?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
      Exit Sub
   End If

   'Elimina Registro de Fecha Limite
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE "
   g_str_Parame = g_str_Parame & "  FROM PRY_DATGENRP "
   g_str_Parame = g_str_Parame & " WHERE DATGENRP_CODPRY = '" & moddat_g_str_Codigo & "' "
   g_str_Parame = g_str_Parame & "   AND DATGENRP_FECREG = '" & Format(grdLstRRPP.TextMatrix(grdLstRRPP.Row, 0), "YYYYMMDD") & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      moddat_g_int_CntErr = 2
      Exit Sub
   End If

   cmd_Salida_Click
End Sub

Private Sub cmd_Grabar_Click()
   If Not IsDate(ipp_FecInfInm.Text) Then
      MsgBox "Debe ingresar Fecha Correcta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecInfInm)
      Exit Sub
   End If
   If Len(Trim(txt_ComenRP.Text)) = 0 Then
      MsgBox "Debe ingresar Comentarios del Registro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ComenRP)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If l_str_Mant = "N" Then
      'Verificar fecha ingresada no sea menor a la anterior
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT DATGENRP_FECREG "
      g_str_Parame = g_str_Parame & "  FROM PRY_DATGENRP "
      g_str_Parame = g_str_Parame & " WHERE DATGENRP_CODPRY = '" & moddat_g_str_Codigo & "' "
      g_str_Parame = g_str_Parame & " ORDER BY DATGENRP_FECREG DESC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         moddat_g_int_CntErr = 2
         Exit Sub
      End If

      If Not (g_rst_Princi.EOF And g_rst_Princi.BOF) Then
         g_rst_Princi.MoveFirst
         If g_rst_Princi!DATGENRP_FECREG > CLng(Format(ipp_FecInfInm.Text, "YYYYMMDD")) Then
            MsgBox "Fecha Limite no puede ser menor a las registradas.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_FecInfInm)
            Exit Sub
         End If
      End If
      
      
      'Buscando Fecha Limite ya registrada
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS CONTADOR "
      g_str_Parame = g_str_Parame & "  FROM PRY_DATGENRP "
      g_str_Parame = g_str_Parame & " WHERE DATGENRP_CODPRY = '" & moddat_g_str_Codigo & "' "
      g_str_Parame = g_str_Parame & "   AND DATGENRP_FECREG = '" & Format(ipp_FecInfInm.Text, "yyyymmdd") & "' "
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         moddat_g_int_CntErr = 2
         Exit Sub
      End If
      
      If Not (g_rst_Princi.EOF And g_rst_Princi.BOF) Then
         g_rst_Princi.MoveFirst
         If g_rst_Princi!CONTADOR > 0 Then
            MsgBox "Fecha Limite ya fue registrada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_FecInfInm)
            Exit Sub
         End If
      End If
   End If

   
   Select Case l_str_Mant
      'Insertar nuevo registro
      Case "N"
            g_str_Parame = "INSERT INTO PRY_DATGENRP ("
            g_str_Parame = g_str_Parame & "DATGENRP_CODPRY, "
            g_str_Parame = g_str_Parame & "DATGENRP_FECREG, "
            g_str_Parame = g_str_Parame & "DATGENRP_COMENT, "
            g_str_Parame = g_str_Parame & "SEGUSUCRE, "
            g_str_Parame = g_str_Parame & "SEGFECCRE, "
            g_str_Parame = g_str_Parame & "SEGHORCRE, "
            g_str_Parame = g_str_Parame & "SEGPLTCRE, "
            g_str_Parame = g_str_Parame & "SEGTERCRE, "
            g_str_Parame = g_str_Parame & "SEGSUCCRE, "
            g_str_Parame = g_str_Parame & "SEGUSUACT, "
            g_str_Parame = g_str_Parame & "SEGFECACT, "
            g_str_Parame = g_str_Parame & "SEGHORACT, "
            g_str_Parame = g_str_Parame & "SEGPLTACT, "
            g_str_Parame = g_str_Parame & "SEGTERACT, "
            g_str_Parame = g_str_Parame & "SEGSUCACT) "
            g_str_Parame = g_str_Parame & "VALUES ( "
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
            g_str_Parame = g_str_Parame & "'" & Format(ipp_FecInfInm.Text, "YYYYMMDD") & "', "
            g_str_Parame = g_str_Parame & "'" & Replace(Trim(txt_ComenRP.Text), "'", "''") & "', "
            
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & Format(date, "YYYYMMDD") & "', "
            g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & Format(date, "YYYYMMDD") & "', "
            g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
      'Modificar registro
      Case "M"
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "UPDATE PRY_DATGENRP SET "
            g_str_Parame = g_str_Parame & " DATGENRP_COMENT='" & Replace(Trim(txt_ComenRP.Text), "'", "''") & "', "
            g_str_Parame = g_str_Parame & " SEGUSUACT='" & modgen_g_str_CodUsu & "',"
            g_str_Parame = g_str_Parame & " SEGFECACT='" & Format(date, "YYYYMMDD") & "',"
            g_str_Parame = g_str_Parame & " SEGHORACT='" & Format(Time, "HHMMSS") & "',"
            g_str_Parame = g_str_Parame & " SEGPLTACT='" & UCase(App.EXEName) & "',"
            g_str_Parame = g_str_Parame & " SEGTERACT='" & modgen_g_str_NombPC & "',"
            g_str_Parame = g_str_Parame & " SEGSUCACT='" & modgen_g_str_CodSuc & "' "
            g_str_Parame = g_str_Parame & "WHERE DATGENRP_CODPRY = '" & moddat_g_str_Codigo & "' AND DATGENRP_FECREG='" & Format(ipp_FecInfInm.Text, "YYYYMMDD") & "'"
   End Select
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If

   cmd_Salida_Click
End Sub

Private Sub cmd_Modifica_Click()
   l_str_Mant = "M"
   
   cmd_Agrega.Enabled = False
   cmd_Modifica.Enabled = False
   cmd_Elimina.Enabled = False
   cmd_Grabar.Enabled = True
   
   ipp_FecInfInm.Enabled = False
   txt_ComenRP.Enabled = True
   ipp_FecInfInm.Text = grdLstRRPP.TextMatrix(grdLstRRPP.Row, 0)
   txt_ComenRP.Text = grdLstRRPP.TextMatrix(grdLstRRPP.Row, 1)
   
   Call gs_SetFocus(txt_ComenRP)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub
 
Private Sub Form_Load()
Dim r_int_Fila As Integer

   Me.Caption = modgen_g_str_NomPlt

   If moddat_g_int_TipCli = 1 Then
      pnl_TituloForm.Caption = "Gestión de Proyectos Vinculados"
   Else
      pnl_TituloForm.Caption = "Gestión de Proyectos No Vinculados"
   End If

   'Buscando Fecha Limite ya registrada
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS CONTADOR "
   g_str_Parame = g_str_Parame & "  FROM PRY_DATGENRP "
   g_str_Parame = g_str_Parame & " WHERE DATGENRP_CODPRY = '" & moddat_g_str_Codigo & "' "
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      moddat_g_int_CntErr = 2
      Exit Sub
   End If
   
   If Not (g_rst_Princi.EOF And g_rst_Princi.BOF) Then
      g_rst_Princi.MoveFirst
      If g_rst_Princi!CONTADOR = 0 Then
         cmd_Modifica.Enabled = False
         cmd_Elimina.Enabled = False
         grdLstRRPP.Rows = 2
      ElseIf g_rst_Princi!CONTADOR > 0 Then
         grdLstRRPP.Rows = 1
      End If
   End If

   cmd_Agrega.Enabled = True
   cmd_Grabar.Enabled = False
   
   ipp_FecInfInm.Enabled = False
   txt_ComenRP.Enabled = False
   
   pnl_Proyecto.Caption = frm_PryNvi_02.txt_NomPry.Text

   'Lista de RRPP
   With grdLstRRPP
      .TextMatrix(0, 0) = "Fecha Limite"
      .TextMatrix(0, 1) = "Comentarios"
      .FixedAlignment(0) = flexAlignCenterCenter
      .FixedAlignment(1) = flexAlignCenterCenter
      .ColAlignment(0) = flexAlignCenterCenter
      .ColWidth(0) = 1200
      .ColWidth(1) = 7400
   End With

   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DATGENRP_FECREG,DATGENRP_COMENT "
   g_str_Parame = g_str_Parame & "  FROM PRY_DATGENRP "
   g_str_Parame = g_str_Parame & " WHERE DATGENRP_CODPRY = '" & moddat_g_str_Codigo & "'"
   g_str_Parame = g_str_Parame & " ORDER BY DATGENRP_CODPRY,DATGENRP_FECREG DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_Fila = 1
  
   Do While Not g_rst_Princi.EOF
      grdLstRRPP.Rows = grdLstRRPP.Rows + 1
      
      grdLstRRPP.TextMatrix(r_int_Fila, 0) = gf_FormatoFecha(CStr(g_rst_Princi!DATGENRP_FECREG))
      grdLstRRPP.TextMatrix(r_int_Fila, 1) = g_rst_Princi!DATGENRP_COMENT
      
      r_int_Fila = r_int_Fila + 1
            
      g_rst_Princi.MoveNext
   Loop
End Sub

Private Sub grdLstRRPP_SelChange()
   If grdLstRRPP.Rows > 2 Then
      grdLstRRPP.RowSel = grdLstRRPP.Row
   End If
End Sub

Private Sub ipp_FecInfInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ComenRP)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "/")
   End If
End Sub

Private Sub txt_ComenRP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If KeyAscii = 13 Then KeyAscii = 0
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "'-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub
