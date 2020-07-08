VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{13E51000-A52B-11D0-86DA-00608CB9FBFB}#5.0#0"; "VCF15.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm_RptSeg_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   1830
   ClientTop       =   2280
   ClientWidth     =   12195
   Icon            =   "AteCli_frm_050.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12195
      _Version        =   65536
      _ExtentX        =   21511
      _ExtentY        =   15849
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   795
         Left            =   30
         TabIndex        =   1
         Top             =   8130
         Width           =   12075
         _Version        =   65536
         _ExtentX        =   21299
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   11340
            Picture         =   "AteCli_frm_050.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   9930
            Picture         =   "AteCli_frm_050.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Excel 
            Height          =   675
            Left            =   10620
            Picture         =   "AteCli_frm_050.frx":0890
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin MSComDlg.CommonDialog CmDlg_Grabar 
            Left            =   60
            Top             =   180
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   7335
         Left            =   60
         TabIndex        =   5
         Top             =   750
         Width           =   12075
         _Version        =   65536
         _ExtentX        =   21299
         _ExtentY        =   12938
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
         Begin VCF150Ctl.F1Book f1_Imprim 
            Height          =   7215
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   12726
            _0              =   $"AteCli_frm_050.frx":0B9A
            _1              =   $"AteCli_frm_050.frx":0FA3
            _2              =   $"AteCli_frm_050.frx":13AC
            _3              =   $"AteCli_frm_050.frx":17B5
            _4              =   $"AteCli_frm_050.frx":1BBE
            _5              =   $"AteCli_frm_050.frx":1FC7
            _6              =   $"AteCli_frm_050.frx":23D1
            _7              =   $"AteCli_frm_050.frx":27DA
            _8              =   $"AteCli_frm_050.frx":2BE3
            _count          =   9
            _ver            =   2
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   12075
         _Version        =   65536
         _ExtentX        =   21299
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
            Left            =   720
            TabIndex        =   8
            Top             =   60
            Width           =   6765
            _Version        =   65536
            _ExtentX        =   11933
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Imprimir Orden de Trabajo - Evaluación de Seguros"
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
            Picture         =   "AteCli_frm_050.frx":2F67
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_RptSeg_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_EmpSeg     As String
Dim l_str_NumEva     As String
Dim l_int_CygTDo     As Integer
Dim l_str_CygNDo     As String
Dim l_str_CygNom     As String
Dim l_str_CygNac     As String
Dim l_str_FecNac     As String
Dim l_str_TitTel     As String
Dim l_str_TitCel     As String
Dim l_str_TitLab     As String
Dim l_str_TitCLb     As String
Dim l_str_CygCel     As String
Dim l_str_CygLab     As String
Dim l_str_CygCLb     As String

Private Sub cmd_Excel_Click()
   Dim r_str_NomArc     As String
   
   On Error GoTo cmd_Excel_Error
   Screen.MousePointer = 11
   
   r_str_NomArc = ""
   
   CmDlg_Grabar.CancelError = True
   CmDlg_Grabar.DialogTitle = "Exportar a Excel"
   CmDlg_Grabar.Filter = "Archivo (*.XLS)"
   CmDlg_Grabar.FileName = "FSEGUR-01.XLS"
   
   CmDlg_Grabar.ShowSave
   
   If Trim$(CmDlg_Grabar.FileName) <> "" Then
      r_str_NomArc = Trim$(CmDlg_Grabar.FileName)
      
      If UCase$(Right$(r_str_NomArc, 4)) <> ".XLS" Then
         r_str_NomArc = r_str_NomArc + ".XLS"
      End If
      
      If Dir(r_str_NomArc) <> "" Then
         If MsgBox("¿ Desea sobreescribir el archivo : " + UCase$(r_str_NomArc), vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
            f1_Imprim.Write UCase$(r_str_NomArc), 4 '***--- Sobreescribe en formato Excel ---***
         End If
      Else
         f1_Imprim.Write UCase$(r_str_NomArc), 4 '***--- Graba en formato Excel ---***
      End If
   End If
   
   Screen.MousePointer = 0
   Exit Sub
   
cmd_Excel_Error:
   Screen.MousePointer = 0
   Exit Sub
End Sub

Private Sub cmd_Imprim_Click()
   If MsgBox("¿Desea imprimir la Orden de Trabajo?", vbYesNo + vbQuestion + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      f1_Imprim.SetSelection 1, 1, f1_Imprim.MaxRow, f1_Imprim.MaxCol
      f1_Imprim.SetPrintAreaFromSelection
      f1_Imprim.PrintLandscape = False
      f1_Imprim.PrintScale = 300
      Call f1_Imprim.FilePrint(True)
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   Call fs_Buscar_InfSeg
   Call fs_Buscar_DatCyg
   
   Call fs_Imprim
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Imprim()
   f1_Imprim.Sheet = 1
   
   'Fecha de Emisión
   f1_Imprim.Row = 3:   f1_Imprim.Col = 13
   f1_Imprim.Text = Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy")
   
   'Número de Solicitud
   f1_Imprim.Row = 5:   f1_Imprim.Col = 11
   f1_Imprim.Text = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4) & " / " & l_str_NumEva
   
   'Nombre de Producto
   f1_Imprim.Row = 11:  f1_Imprim.Col = 4
   f1_Imprim.Text = moddat_g_str_NomPrd
   
   'Descripción de Modalidad
   f1_Imprim.Row = 12:  f1_Imprim.Col = 4
   f1_Imprim.Text = moddat_g_str_DesMod
   
   'Empresa Aseguradora
   f1_Imprim.Row = 13:  f1_Imprim.Col = 4
   f1_Imprim.Text = l_str_EmpSeg
   
   
   'Nombre Cliente
   f1_Imprim.Row = 18:  f1_Imprim.Col = 4
   f1_Imprim.Text = moddat_g_str_NomCli
   
   f1_Imprim.Row = 19:  f1_Imprim.Col = 4
   f1_Imprim.Text = moddat_gf_Consulta_Pardes("203", CStr(moddat_g_int_TipDoc)) & "-" & Trim(moddat_g_str_NumDoc)
   
   f1_Imprim.Row = 19:  f1_Imprim.Col = 10
   f1_Imprim.Text = l_str_FecNac
   
   f1_Imprim.Row = 20:  f1_Imprim.Col = 4
   f1_Imprim.Text = l_str_TitTel
   
   f1_Imprim.Row = 21:  f1_Imprim.Col = 4
   f1_Imprim.Text = l_str_TitCel
   
   f1_Imprim.Row = 21:  f1_Imprim.Col = 6
   f1_Imprim.Text = l_str_TitCLb
   
   f1_Imprim.Row = 21:  f1_Imprim.Col = 10
   f1_Imprim.Text = l_str_TitLab
   
   
   'Nombre Cónyuge
   If l_int_CygTDo > 0 Then
      f1_Imprim.Row = 22:  f1_Imprim.Col = 4
      f1_Imprim.Text = l_str_CygNom
      
      f1_Imprim.Row = 23:  f1_Imprim.Col = 4
      f1_Imprim.Text = moddat_gf_Consulta_Pardes("203", CStr(l_int_CygTDo)) & "-" & Trim(l_str_CygNom)
      
      f1_Imprim.Row = 23:  f1_Imprim.Col = 10
      f1_Imprim.Text = l_str_CygNac
      
      f1_Imprim.Row = 24:  f1_Imprim.Col = 4
      f1_Imprim.Text = l_str_CygCel
      
      f1_Imprim.Row = 24:  f1_Imprim.Col = 6
      f1_Imprim.Text = l_str_CygCLb
      
      f1_Imprim.Row = 24:  f1_Imprim.Col = 10
      f1_Imprim.Text = l_str_CygLab
   End If
End Sub

Private Sub fs_Buscar_DatCyg()
   Dim r_str_FecOcu  As String
   Dim l_rst_Genera  As ADODB.Recordset

   l_str_TitTel = ""
   l_str_TitCel = ""
   l_str_TitLab = ""
   l_str_TitCLb = ""
   l_str_FecNac = ""
   l_int_CygTDo = 0
   l_str_CygNDo = ""
   l_str_CygNom = ""
   l_str_CygNac = ""
   l_str_CygCel = ""
   l_str_CygLab = ""
   l_str_CygCLb = ""

   'Obteniendo ID de Cliente Titular
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   g_rst_Princi.MoveFirst

   l_int_CygTDo = g_rst_Princi!DatGen_CygTDo
   l_str_CygNDo = Trim(g_rst_Princi!DatGen_CygNDo)
   
   r_str_FecOcu = Format(g_rst_Princi!DATGEN_NACFEC, "00000000")
   r_str_FecOcu = Right(r_str_FecOcu, 2) & "/" & Mid(r_str_FecOcu, 5, 2) & "/" & Left(r_str_FecOcu, 4)
   l_str_FecNac = r_str_FecOcu

   l_str_TitTel = Trim(g_rst_Princi!DatGen_Telefo)
   l_str_TitCel = Trim(g_rst_Princi!DatGen_NUMCEL)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Obteniendo Teléfonos Laborales
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = 1"

   If gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      g_rst_Princi.MoveFirst
   
      Select Case g_rst_Princi!ActEco_CodAct
         Case 11, 12, 31, 41
            g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
            g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_TipDoc) & " AND "
            g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_NumDoc) & "' "
         
            If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
               Exit Sub
            End If
      
            l_rst_Genera.MoveFirst
         
            If g_rst_Princi!ActEco_CodAct = 21 Or g_rst_Princi!ActEco_CodAct = 31 Then
               l_str_TitLab = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
         
               If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
                  l_str_TitLab = l_str_TitLab & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
               End If
            Else
               If Len(Trim(g_rst_Princi!ActEco_Sucurs & "")) > 0 Then
                  l_str_TitLab = Trim(g_rst_Princi!ActEco_Telef1 & "")
            
                  If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
                     l_str_TitLab = l_str_TitLab & Trim(g_rst_Princi!ActEco_Telef2 & "")
                  End If
                  
                  If Len(Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")) > 0 Then
                     l_str_TitLab = l_str_TitLab & " / Anx: " & Trim(g_rst_Princi!ActEco_Dep_NumAnx)
                  End If
               Else
                  l_str_TitLab = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
         
                  If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
                     l_str_TitLab = l_str_TitLab & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
                  End If
                  
                  If Len(Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")) > 0 Then
                     l_str_TitLab = l_str_TitLab & " / Anx: " & Trim(g_rst_Princi!ActEco_Dep_NumAnx)
                  End If
               End If
               
               If Len(Trim(g_rst_Princi!ActEco_Dep_Celula & "")) > 0 Then
                  l_str_TitCLb = Trim(g_rst_Princi!ActEco_Dep_Celula)
               End If
            End If
         
         Case 21
            If g_rst_Princi!ActEco_Ind_ConLoc = 2 Then
               l_str_TitLab = Trim(g_rst_Princi!ActEco_Telef1 & "")
         
               If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
                  l_str_TitLab = l_str_TitLab & Trim(g_rst_Princi!ActEco_Telef2 & "")
               End If
            Else
               g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
               g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TDoEmp) & " AND "
               g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_Ind_NDoEmp) & "' "
      
               If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
                  Exit Sub
               End If
   
               l_rst_Genera.MoveFirst
               
               l_str_TitLab = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
            
               If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
                  l_str_TitLab = l_str_TitLab & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
               End If
            
               l_rst_Genera.Close
               Set l_rst_Genera = Nothing
            End If
      End Select
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing


   'Información del Cónyuge
   If l_int_CygTDo > 0 Then
      'Obteniendo Nombre de Cónyuge
      g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(l_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & l_str_CygNDo & "' "
   
      If gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         g_rst_Princi.MoveFirst
         
         l_str_CygNom = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & " " & Trim(g_rst_Princi!DatGen_Nombre)
      
         r_str_FecOcu = Format(g_rst_Princi!DATGEN_NACFEC, "00000000")
         r_str_FecOcu = Right(r_str_FecOcu, 2) & "/" & Mid(r_str_FecOcu, 5, 2) & "/" & Left(r_str_FecOcu, 4)
         l_str_CygNac = r_str_FecOcu
      
         l_str_CygCel = Trim(g_rst_Princi!DatGen_NUMCEL)
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      'Obteniendo Teléfonos Laborales
      g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
      g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(l_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & l_str_CygNDo & "' AND "
      g_str_Parame = g_str_Parame & "ACTECO_ORDACT = 1"
   
      If gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         g_rst_Princi.MoveFirst
      
         Select Case g_rst_Princi!ActEco_CodAct
            Case 11, 12, 31, 41
               g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
               g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_TipDoc) & " AND "
               g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_NumDoc) & "' "
            
               If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
                  Exit Sub
               End If
         
               l_rst_Genera.MoveFirst
            
               If g_rst_Princi!ActEco_CodAct = 21 Or g_rst_Princi!ActEco_CodAct = 31 Then
                  l_str_CygLab = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
            
                  If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
                     l_str_CygLab = l_str_CygLab & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
                  End If
               Else
                  If Len(Trim(g_rst_Princi!ActEco_Sucurs & "")) > 0 Then
                     l_str_CygLab = Trim(g_rst_Princi!ActEco_Telef1 & "")
               
                     If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
                        l_str_CygLab = l_str_CygLab & Trim(g_rst_Princi!ActEco_Telef2 & "")
                     End If
                     
                     If Len(Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")) > 0 Then
                        l_str_CygLab = l_str_CygLab & " / Anx: " & Trim(g_rst_Princi!ActEco_Dep_NumAnx)
                     End If
                  Else
                     l_str_CygLab = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
            
                     If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
                        l_str_CygLab = l_str_CygLab & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
                     End If
                     
                     If Len(Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")) > 0 Then
                        l_str_CygLab = l_str_CygLab & " / Anx: " & Trim(g_rst_Princi!ActEco_Dep_NumAnx)
                     End If
                  End If
                  
                  If Len(Trim(g_rst_Princi!ActEco_Dep_Celula & "")) > 0 Then
                     l_str_CygCLb = Trim(g_rst_Princi!ActEco_Dep_Celula)
                  End If
               End If
            
            Case 21
               If g_rst_Princi!ActEco_Ind_ConLoc = 2 Then
                  l_str_CygLab = Trim(g_rst_Princi!ActEco_Telef1 & "")
            
                  If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
                     l_str_CygLab = l_str_CygLab & Trim(g_rst_Princi!ActEco_Telef2 & "")
                  End If
               Else
                  g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
                  g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TDoEmp) & " AND "
                  g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_Ind_NDoEmp) & "' "
         
                  If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
                     Exit Sub
                  End If
      
                  l_rst_Genera.MoveFirst
                  
                  l_str_CygLab = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
               
                  If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
                     l_str_CygLab = l_str_CygLab & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
                  End If
               
                  l_rst_Genera.Close
                  Set l_rst_Genera = Nothing
               End If
         End Select
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
End Sub

Private Sub fs_Buscar_InfSeg()
   l_str_EmpSeg = ""
   l_str_NumEva = ""
   
   g_str_Parame = "SELECT * FROM TRA_EVASEG WHERE "
   g_str_Parame = g_str_Parame & "EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY EVASEG_NUMEVA DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Empresa de Seguros
   l_str_EmpSeg = moddat_gf_Consulta_Pardes("508", Format(g_rst_Princi!SOLSEG_CODEMP, "000000"))
   l_str_NumEva = CStr(g_rst_Princi!SOLSEG_NUMEVA)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub


