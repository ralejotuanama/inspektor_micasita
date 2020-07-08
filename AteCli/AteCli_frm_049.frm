VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{13E51000-A52B-11D0-86DA-00608CB9FBFB}#5.0#0"; "VCF15.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm_RptTas_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   1605
   ClientTop       =   1245
   ClientWidth     =   12165
   Icon            =   "AteCli_frm_049.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12165
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
         TabIndex        =   3
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
         Begin VB.CommandButton cmd_Excel 
            Height          =   675
            Left            =   10620
            Picture         =   "AteCli_frm_049.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   9930
            Picture         =   "AteCli_frm_049.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   11340
            Picture         =   "AteCli_frm_049.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
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
         TabIndex        =   1
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
            TabIndex        =   2
            Top             =   60
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   12726
            _0              =   $"AteCli_frm_049.frx":0B9A
            _1              =   $"AteCli_frm_049.frx":0FA3
            _2              =   $"AteCli_frm_049.frx":13AC
            _3              =   $"AteCli_frm_049.frx":17B5
            _4              =   $"AteCli_frm_049.frx":1BBE
            _5              =   $"AteCli_frm_049.frx":1FC7
            _6              =   $"AteCli_frm_049.frx":23D0
            _7              =   $"AteCli_frm_049.frx":27D8
            _8              =   $"AteCli_frm_049.frx":2BE1
            _9              =   ")I)@lt-@@@@@@F@,8B3F"
            _count          =   10
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
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Imprimir Orden de Trabajo - Tasación"
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
            Picture         =   "AteCli_frm_049.frx":2FEA
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_RptTas_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_Direc1     As String
Dim l_str_Direc2     As String
Dim l_str_Direc3     As String
Dim l_str_EmpPer     As String
Dim l_str_NumEva     As String
Dim l_str_NomPr1     As String
Dim l_str_DoiPr1     As String
Dim l_str_NomPr2     As String
Dim l_str_DoiPr2     As String
Dim l_int_CygTDo     As Integer
Dim l_str_CygNDo     As String
Dim l_str_CygNom     As String

Private Sub cmd_Excel_Click()
   Dim r_str_NomArc     As String
   
   On Error GoTo cmd_Excel_Error
   Screen.MousePointer = 11
   
   r_str_NomArc = ""
   
   CmDlg_Grabar.CancelError = True
   CmDlg_Grabar.DialogTitle = "Exportar a Excel"
   CmDlg_Grabar.Filter = "Archivo (*.XLS)"
   CmDlg_Grabar.FileName = "FTASAC-01.XLS"
   
   CmDlg_Grabar.ShowSave
   
   If Trim(CmDlg_Grabar.FileName) <> "" Then
      r_str_NomArc = Trim(CmDlg_Grabar.FileName)
      
      If UCase(Right(r_str_NomArc, 4)) <> ".XLS" Then
         r_str_NomArc = r_str_NomArc + ".XLS"
      End If
      
      If Dir(r_str_NomArc) <> "" Then
         If MsgBox("¿ Desea sobreescribir el archivo : " + UCase(r_str_NomArc), vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
            f1_Imprim.Write UCase(r_str_NomArc), 4 '***--- Sobreescribe en formato Excel ---***
         End If
      Else
         f1_Imprim.Write UCase(r_str_NomArc), 4 '***--- Graba en formato Excel ---***
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

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Carga_InfTas
   Call fs_Carga_DatInm
   Call fs_Carga_DatCyg
   
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
   
   'Empresa de Peritaje
   f1_Imprim.Row = 13:  f1_Imprim.Col = 4
   f1_Imprim.Text = l_str_EmpPer
   
   'Dirección de Inmueble
   f1_Imprim.Row = 23:  f1_Imprim.Col = 4
   f1_Imprim.Text = l_str_Direc1
   
   f1_Imprim.Row = 24:  f1_Imprim.Col = 4
   f1_Imprim.Text = l_str_Direc2
   
   f1_Imprim.Row = 25:  f1_Imprim.Col = 4
   f1_Imprim.Text = l_str_Direc3
   
   'Propietario de Inmueble
   f1_Imprim.Row = 31:  f1_Imprim.Col = 4
   f1_Imprim.Text = l_str_NomPr1
   
   f1_Imprim.Row = 31:  f1_Imprim.Col = 11
   f1_Imprim.Text = l_str_DoiPr1
   
   f1_Imprim.Row = 32:  f1_Imprim.Col = 4
   f1_Imprim.Text = l_str_NomPr2

   f1_Imprim.Row = 32:  f1_Imprim.Col = 11
   f1_Imprim.Text = l_str_DoiPr2
   
   'Nombre Cliente
   f1_Imprim.Row = 29:  f1_Imprim.Col = 4
   f1_Imprim.Text = moddat_g_str_NomCli
   
   f1_Imprim.Row = 29:  f1_Imprim.Col = 11
   f1_Imprim.Text = moddat_gf_Consulta_Pardes("203", CStr(moddat_g_int_TipDoc)) & "-" & Trim(moddat_g_str_NumDoc)
   
   'Nombre Cónyuge
   If l_int_CygTDo > 0 Then
      f1_Imprim.Row = 30:  f1_Imprim.Col = 4
      f1_Imprim.Text = l_str_CygNom
   
      f1_Imprim.Row = 30:  f1_Imprim.Col = 11
      f1_Imprim.Text = moddat_gf_Consulta_Pardes("203", CStr(l_int_CygTDo)) & "-" & Trim(l_str_CygNDo)
   End If
End Sub

Private Sub fs_Carga_InfTas()
   l_str_EmpPer = ""
   l_str_NumEva = ""
   
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY EVATAS_NUMEVA DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Empresa de Peritaje
   l_str_EmpPer = moddat_gf_Consulta_Pardes("507", Format(g_rst_Princi!EVATAS_CODEMP, "000000"))
   
   l_str_NumEva = CStr(g_rst_Princi!EVATAS_NUMEVA) & " "
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Carga_DatInm()
   Dim r_str_TipVia  As String
   Dim r_str_TipZon  As String
   Dim r_str_Depart  As String
   Dim r_str_Provin  As String
   Dim r_str_Distri  As String
   
   l_str_Direc1 = ""
   l_str_Direc2 = ""
   l_str_Direc3 = ""
   l_str_DoiPr1 = ""
   l_str_NomPr1 = ""
   l_str_DoiPr2 = ""
   l_str_NomPr2 = ""
   
   
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SOLINM_SITUAC = 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   r_str_TipVia = moddat_gf_Consulta_Pardes("201", CStr(g_rst_Princi!SOLINM_TIPVIA))
   r_str_TipZon = moddat_gf_Consulta_Pardes("202", CStr(g_rst_Princi!SOLINM_TIPZON))

   l_str_Direc1 = r_str_TipVia & " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMERO)
   
   If Len(Trim(Trim(g_rst_Princi!SOLINM_INTDPT))) > 0 Then
      l_str_Direc1 = l_str_Direc1 & " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")"
   End If
   
   If Len(Trim(Trim(g_rst_Princi!SOLINM_NOMZON))) > 0 Then
      l_str_Direc1 = l_str_Direc1 & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!SOLINM_NOMZON) & Chr(13) & Chr(10)
   End If
   
   'Departamento
   r_str_Depart = moddat_gf_Consulta_Pardes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000")
   
   'Provincia
   r_str_Provin = moddat_gf_Consulta_Pardes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00")
   
   'Distrito
   r_str_Distri = moddat_gf_Consulta_Pardes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
   
   l_str_Direc2 = l_str_Direc2 & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   
   'Referencia
   If Len(Trim(g_rst_Princi!SOLINM_REFERE)) > 0 Then
      l_str_Direc3 = Trim(g_rst_Princi!SOLINM_REFERE)
   End If

   'Información del Propietario
   If g_rst_Princi!SOLINM_TIPPER = 2 Then
      'Persona Jurídica
      l_str_DoiPr1 = moddat_gf_Consulta_Pardes("203", CStr(g_rst_Princi!SOLINM_PROTDO)) & "-" & Trim(g_rst_Princi!SOLINM_PRONDO)
      l_str_NomPr1 = Trim(g_rst_Princi!SOLINM_PRORZS)
   Else
      'Persona Natural
      l_str_DoiPr1 = moddat_gf_Consulta_Pardes("203", CStr(g_rst_Princi!SOLINM_PROTDO)) & "-" & Trim(g_rst_Princi!SOLINM_PRONDO)
      l_str_NomPr1 = Trim(g_rst_Princi!SOLINM_PROAPP) & " " & Trim(g_rst_Princi!SOLINM_PROAPM) & " " & Trim(g_rst_Princi!SOLINM_PRONOM)
      
      If g_rst_Princi!SOLINM_CYGTDO > 0 Then
         l_str_DoiPr2 = moddat_gf_Consulta_Pardes("203", CStr(g_rst_Princi!SOLINM_CYGTDO)) & "-" & Trim(g_rst_Princi!SOLINM_CYGNDO)
         l_str_NomPr2 = Trim(g_rst_Princi!SOLINM_CYGAPP) & " " & Trim(g_rst_Princi!SOLINM_CYGAPM) & " " & Trim(g_rst_Princi!SOLINM_CYGNOM)
      End If
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Carga_DatCyg()
   l_int_CygTDo = 0
   l_str_CygNDo = ""
   l_str_CygNom = ""
   
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

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   If l_int_CygTDo > 0 Then
      'Obteniendo Nombre de Cónyuge
      g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(l_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & l_str_CygNDo & "' "
   
      If gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         l_str_CygNom = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & " " & Trim(g_rst_Princi!DatGen_Nombre)
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
End Sub

