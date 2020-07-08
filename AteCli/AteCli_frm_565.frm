VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_RptSol_41 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13875
   Icon            =   "AteCli_frm_565.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   13875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel20 
      Height          =   7200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13875
      _Version        =   65536
      _ExtentX        =   24474
      _ExtentY        =   12700
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   45
         TabIndex        =   1
         Top             =   60
         Width           =   13755
         _Version        =   65536
         _ExtentX        =   24262
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   300
            Left            =   660
            TabIndex        =   2
            Top             =   180
            Width           =   5520
            _Version        =   65536
            _ExtentX        =   9737
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Reporte Detalle de Desembolsos Mensuales"
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
            Left            =   105
            Picture         =   "AteCli_frm_565.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel19 
         Height          =   645
         Left            =   45
         TabIndex        =   3
         Top             =   765
         Width           =   13755
         _Version        =   65536
         _ExtentX        =   24262
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
            Left            =   13140
            Picture         =   "AteCli_frm_565.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   45
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExcRes 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_565.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel - Resumido"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5730
         Left            =   15
         TabIndex        =   6
         Top             =   1425
         Width           =   13800
         _Version        =   65536
         _ExtentX        =   24342
         _ExtentY        =   10107
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_LisDes 
            Height          =   4830
            Left            =   45
            TabIndex        =   7
            Top             =   870
            Width           =   13650
            _ExtentX        =   24077
            _ExtentY        =   8520
            _Version        =   393216
            Rows            =   16
            Cols            =   16
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   7485
            TabIndex        =   8
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "FEB"
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   7005
            TabIndex        =   9
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ENE"
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
            Left            =   9405
            TabIndex        =   10
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "JUN"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   8445
            TabIndex        =   11
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ABR"
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   285
            Left            =   9885
            TabIndex        =   12
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "JUL"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   7965
            TabIndex        =   13
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "MAR"
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   8925
            TabIndex        =   14
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "MAY"
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
            Left            =   11805
            TabIndex        =   15
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "NOV"
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   10845
            TabIndex        =   16
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "SET"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   12285
            TabIndex        =   17
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "DIC"
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   10365
            TabIndex        =   18
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "AGO"
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
         Begin Threed.SSPanel SSPanel17 
            Height          =   285
            Left            =   11325
            TabIndex        =   19
            Top             =   585
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "OCT"
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
         Begin Threed.SSPanel SSPanel18 
            Height          =   285
            Left            =   12765
            TabIndex        =   20
            Top             =   585
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "TOTAL"
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
         Begin Threed.SSPanel SSPanel1 
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   585
            Width           =   4335
            _Version        =   65536
            _ExtentX        =   7646
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Proyecto"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   4380
            TabIndex        =   22
            Top             =   585
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Vinculado"
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
            Left            =   5370
            TabIndex        =   23
            Top             =   585
            Width           =   1650
            _Version        =   65536
            _ExtentX        =   2910
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Modalidad"
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
         Begin VB.Label lbl_NomCon 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   240
            Width           =   6555
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ExpExcRes_Click()
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExcRes
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(grd_LisDes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_LisDes.ColWidth(0) = 4300   ' Proyecto
   grd_LisDes.ColWidth(1) = 980    ' Vinculado
   grd_LisDes.ColAlignment(1) = flexAlignCenterCenter
   grd_LisDes.ColWidth(2) = 1650   ' Modalidad
   grd_LisDes.ColAlignment(2) = flexAlignCenterCenter
   grd_LisDes.ColWidth(3) = 470    ' MES 1
   grd_LisDes.ColWidth(4) = 470    ' MES 2
   grd_LisDes.ColWidth(5) = 490    ' MES 3
   grd_LisDes.ColWidth(6) = 490    ' MES 4
   grd_LisDes.ColWidth(7) = 500    ' MES 5
   grd_LisDes.ColWidth(8) = 500    ' MES 6
   grd_LisDes.ColWidth(9) = 470    ' MES 7
   grd_LisDes.ColWidth(10) = 470   ' MES 8
   grd_LisDes.ColWidth(11) = 470   ' MES 9
   grd_LisDes.ColWidth(12) = 500   ' MES 10
   grd_LisDes.ColWidth(13) = 470   ' MES 11
   grd_LisDes.ColWidth(14) = 490   ' MES 12
   grd_LisDes.ColWidth(15) = 600   ' TOTAL
   Call gs_LimpiaGrid(grd_LisDes)
End Sub

Private Sub fs_GenExcRes()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_ConVer        As Integer
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE RESUMEN DE DESEMBOLSOS MENSUALES"
      .Range(.Cells(2, 2), .Cells(2, 6)).Merge
      .Range(.Cells(2, 2), .Cells(2, 6)).Font.Bold = True
      .Range(.Cells(6, 2), .Cells(6, 17)).Font.Name = "Calibri"
      .Range(.Cells(6, 2), .Cells(6, 17)).Font.Size = 10
      
      .Cells(6, 2) = "PROYECTO"
      .Cells(6, 3) = "'" & "VINCULADO"
      .Cells(6, 4) = "'" & "MODALIDAD"
      .Cells(6, 5) = "'" & "ENE"
      .Cells(6, 6) = "'" & "FEB"
      .Cells(6, 7) = "'" & "MAR"
      .Cells(6, 8) = "'" & "ABR"
      .Cells(6, 9) = "'" & "MAY"
      .Cells(6, 10) = "'" & "JUN"
      .Cells(6, 11) = "'" & "JUL"
      .Cells(6, 12) = "'" & "AGO"
      .Cells(6, 13) = "'" & "SET"
      .Cells(6, 14) = "'" & "OCT"
      .Cells(6, 15) = "'" & "NOV"
      .Cells(6, 16) = "'" & "DIC"
      .Cells(6, 17) = "'" & "TOTAL"
      
      .Range(.Cells(6, 2), .Cells(6, 17)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(6, 2), .Cells(6, 17)).Font.Bold = True
      .Range(.Cells(6, 2), .Cells(6, 17)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 30
      .Columns("C").ColumnWidth = 10
      .Columns("D").ColumnWidth = 16
      .Columns("E").ColumnWidth = 5
      .Columns("E").NumberFormat = "###,###,###,##0"
      .Columns("F").ColumnWidth = 5
      .Columns("F").NumberFormat = "###,###,###,##0"
      .Columns("G").ColumnWidth = 5
      .Columns("G").NumberFormat = "###,###,###,##0"
      .Columns("H").ColumnWidth = 5
      .Columns("H").NumberFormat = "###,###,###,##0"
      .Columns("I").ColumnWidth = 5
      .Columns("I").NumberFormat = "###,###,###,##0"
      .Columns("J").ColumnWidth = 5
      .Columns("J").NumberFormat = "###,###,###,##0"
      .Columns("K").ColumnWidth = 5
      .Columns("K").NumberFormat = "###,###,###,##0"
      .Columns("L").ColumnWidth = 5
      .Columns("L").NumberFormat = "###,###,###,##0"
      .Columns("M").ColumnWidth = 5
      .Columns("M").NumberFormat = "###,###,###,##0"
      .Columns("N").ColumnWidth = 5
      .Columns("N").NumberFormat = "###,###,###,##0"
      .Columns("O").ColumnWidth = 5
      .Columns("O").NumberFormat = "###,###,###,##0"
      .Columns("P").ColumnWidth = 5
      .Columns("P").NumberFormat = "###,###,###,##0"
      .Columns("Q").ColumnWidth = 5
      .Columns("Q").NumberFormat = "###,###,###,##0"
      .Columns("R").ColumnWidth = 8
      .Columns("R").NumberFormat = "###,###,###,##0"
      
      r_int_ConVer = 7
      For r_int_Contad = 0 To grd_LisDes.Rows - 1
         If r_int_Contad = 0 Then
            .Range(.Cells(r_int_ConVer, 2), .Cells(r_int_ConVer, 17)).Font.Bold = True
            .Cells(4, 2) = lbl_NomCon.Caption
            .Range(.Cells(4, 2), .Cells(4, 2)).Merge
            .Range(.Cells(4, 2), .Cells(4, 2)).Font.Bold = True
         End If
         
         .Cells(r_int_ConVer, 2) = grd_LisDes.TextMatrix(r_int_Contad, 0)
         .Cells(r_int_ConVer, 3) = grd_LisDes.TextMatrix(r_int_Contad, 1)
         .Range(.Cells(r_int_ConVer, 3), .Cells(r_int_ConVer, 3)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_ConVer, 4) = grd_LisDes.TextMatrix(r_int_Contad, 2)
         .Range(.Cells(r_int_ConVer, 4), .Cells(r_int_ConVer, 4)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_ConVer, 5) = grd_LisDes.TextMatrix(r_int_Contad, 3)
         .Cells(r_int_ConVer, 6) = grd_LisDes.TextMatrix(r_int_Contad, 4)
         .Cells(r_int_ConVer, 7) = grd_LisDes.TextMatrix(r_int_Contad, 5)
         .Cells(r_int_ConVer, 8) = grd_LisDes.TextMatrix(r_int_Contad, 6)
         .Cells(r_int_ConVer, 9) = grd_LisDes.TextMatrix(r_int_Contad, 7)
         .Cells(r_int_ConVer, 10) = grd_LisDes.TextMatrix(r_int_Contad, 8)
         .Cells(r_int_ConVer, 11) = grd_LisDes.TextMatrix(r_int_Contad, 9)
         .Cells(r_int_ConVer, 12) = grd_LisDes.TextMatrix(r_int_Contad, 10)
         .Cells(r_int_ConVer, 13) = grd_LisDes.TextMatrix(r_int_Contad, 11)
         .Cells(r_int_ConVer, 14) = grd_LisDes.TextMatrix(r_int_Contad, 12)
         .Cells(r_int_ConVer, 15) = grd_LisDes.TextMatrix(r_int_Contad, 13)
         .Cells(r_int_ConVer, 16) = grd_LisDes.TextMatrix(r_int_Contad, 14)
         .Cells(r_int_ConVer, 17) = Val(grd_LisDes.TextMatrix(r_int_Contad, 3)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 4)) + _
                                    Val(grd_LisDes.TextMatrix(r_int_Contad, 5)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 6)) + _
                                    Val(grd_LisDes.TextMatrix(r_int_Contad, 7)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 8)) + _
                                    Val(grd_LisDes.TextMatrix(r_int_Contad, 9)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 10)) + _
                                    Val(grd_LisDes.TextMatrix(r_int_Contad, 11)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 12)) + _
                                    Val(grd_LisDes.TextMatrix(r_int_Contad, 13)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 14))
         
         DoEvents
         r_int_ConVer = r_int_ConVer + 1
      Next
      
      r_obj_Excel.Visible = True
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Public Sub LlenarData(Descripcion As String, Mes As String, Anio As String, Tipo As Integer)
Dim r_int_PerMes      As Integer
Dim r_int_PerAno      As Integer
Dim r_int_Tot01       As Integer
Dim r_int_Tot02       As Integer
Dim r_int_Tot03       As Integer
Dim r_int_Tot04       As Integer
Dim r_int_Tot05       As Integer
Dim r_int_Tot06       As Integer
Dim r_int_Tot07       As Integer
Dim r_int_Tot08       As Integer
Dim r_int_Tot09       As Integer
Dim r_int_Tot10       As Integer
Dim r_int_Tot11       As Integer
Dim r_int_Tot12       As Integer
Dim r_int_TotAc       As Integer

   r_int_PerMes = CInt(Mes)
   r_int_PerAno = CInt(Anio)
   
   
    g_str_Parame = ""
    g_str_Parame = g_str_Parame + " SELECT * FROM RPT_TABLA_TEMP "
    g_str_Parame = g_str_Parame + "  WHERE RPT_PERMES = '" & CInt(r_int_PerMes) & "' "
    g_str_Parame = g_str_Parame + "    AND RPT_PERANO = '" & CInt(r_int_PerAno) & "' "
    g_str_Parame = g_str_Parame + "    AND RPT_TERCRE = '" & modgen_g_str_NombPC & " '"
    g_str_Parame = g_str_Parame + "    AND RPT_USUCRE = '" & modgen_g_str_CodUsu & " '"
    If Tipo = 1 Then
        g_str_Parame = g_str_Parame + "    AND RPT_NOMBRE = 'CONSEJEROS' "
    Else
        g_str_Parame = g_str_Parame + "    AND RPT_NOMBRE = 'TIPO_EVALUACION' "
    End If
    g_str_Parame = g_str_Parame + "    AND RPT_MONEDA = 0 "
    g_str_Parame = g_str_Parame + "    AND RPT_DESCRI = '" & Trim(Descripcion) & "' "
    g_str_Parame = g_str_Parame + " ORDER BY RPT_DESCRI, RPT_VALCAD01, RPT_VALNUM13 DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      grd_LisDes.Rows = 0
      If Tipo = 1 Then
        lbl_NomCon.Caption = " Consejero : " & Trim(g_rst_Princi!RPT_DESCRI)
      Else
        lbl_NomCon.Caption = " Tipo de Evaluación : " & Trim(g_rst_Princi!RPT_DESCRI)
      End If
      Do While Not g_rst_Princi.EOF
         If grd_LisDes.Row = 0 Then
            grd_LisDes.Col = 3
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 4
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 5
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 6
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 7
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 8
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 9
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 10
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 11
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 12
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 13
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 14
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
            grd_LisDes.Col = 15
            grd_LisDes.CellFontBold = True
            grd_LisDes.CellFontSize = 10
         End If
         
         grd_LisDes.Rows = grd_LisDes.Rows + 1
         If grd_LisDes.Rows > 1 Then grd_LisDes.Row = grd_LisDes.Row + 1
         
         grd_LisDes.Col = 0
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         grd_LisDes.Text = IIf(IsNull(g_rst_Princi!RPT_VALCAD02), "", g_rst_Princi!RPT_VALCAD02)
                     
         grd_LisDes.Col = 1
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         grd_LisDes.Text = IIf(IsNull(g_rst_Princi!RPT_VALCAD03), "", g_rst_Princi!RPT_VALCAD03)
         
         grd_LisDes.Col = 2
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         grd_LisDes.Text = IIf(IsNull(g_rst_Princi!RPT_VALCAD04), "", g_rst_Princi!RPT_VALCAD04)
   
         grd_LisDes.Col = 3
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot01 = r_int_Tot01 + g_rst_Princi!RPT_VALNUM01
         
         grd_LisDes.Col = 4
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot02 = r_int_Tot02 + g_rst_Princi!RPT_VALNUM02
         
         grd_LisDes.Col = 5
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot03 = r_int_Tot03 + g_rst_Princi!RPT_VALNUM03
   
         grd_LisDes.Col = 6
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM04, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot04 = r_int_Tot04 + g_rst_Princi!RPT_VALNUM04
   
         grd_LisDes.Col = 7
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM05, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot05 = r_int_Tot05 + g_rst_Princi!RPT_VALNUM05
   
         grd_LisDes.Col = 8
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM06, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot06 = r_int_Tot06 + g_rst_Princi!RPT_VALNUM06
   
         grd_LisDes.Col = 9
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM07, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot07 = r_int_Tot07 + g_rst_Princi!RPT_VALNUM07
   
         grd_LisDes.Col = 10
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM08, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot08 = r_int_Tot08 + g_rst_Princi!RPT_VALNUM08
   
         grd_LisDes.Col = 11
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM09, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot09 = r_int_Tot09 + g_rst_Princi!RPT_VALNUM09
   
         grd_LisDes.Col = 12
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM10, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot10 = r_int_Tot10 + g_rst_Princi!RPT_VALNUM10
   
         grd_LisDes.Col = 13
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM11, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot11 = r_int_Tot11 + g_rst_Princi!RPT_VALNUM11
   
         grd_LisDes.Col = 14
         grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM12, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         r_int_Tot12 = r_int_Tot12 + g_rst_Princi!RPT_VALNUM12
         
         r_int_TotAc = r_int_TotAc + g_rst_Princi!RPT_VALNUM01 + g_rst_Princi!RPT_VALNUM02 + g_rst_Princi!RPT_VALNUM03 + _
                                     g_rst_Princi!RPT_VALNUM04 + g_rst_Princi!RPT_VALNUM05 + g_rst_Princi!RPT_VALNUM06 + _
                                     g_rst_Princi!RPT_VALNUM07 + g_rst_Princi!RPT_VALNUM08 + g_rst_Princi!RPT_VALNUM09 + _
                                     g_rst_Princi!RPT_VALNUM10 + g_rst_Princi!RPT_VALNUM11 + g_rst_Princi!RPT_VALNUM12
         
         grd_LisDes.Col = 15
         grd_LisDes.Text = Format(r_int_TotAc, "###,###,##0")
         grd_LisDes.CellFontName = "Arial"
         grd_LisDes.CellFontSize = 8
         
         r_int_TotAc = 0
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_LisDes)
   End If
End Sub

Private Sub grd_LisDes_SelChange()
   If grd_LisDes.Rows > 2 Then
      grd_LisDes.RowSel = grd_LisDes.Row
   End If
End Sub
