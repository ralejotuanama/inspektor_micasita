VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bat_ClsCnx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public g_cnx_Conexion      As Connection
Public Command             As Command

Public Function gf_AbreBaseDatos(p_CadenaConexion As String) As Boolean
   'Realiza la connección a la base de datos SQL
   
   On Error GoTo Error_gf_AbreBaseDatos
   
   Set g_cnx_Conexion = New ADODB.Connection
   g_cnx_Conexion.ConnectionString = p_CadenaConexion
   g_cnx_Conexion.Open
   gf_AbreBaseDatos = True
   Exit Function

Error_gf_AbreBaseDatos:
   gf_AbreBaseDatos = False
   
End Function

Public Function gf_EjecutaComandoSQL(p_Cadena As String, p_RecordSet As Recordset, Opcion As Integer) As Integer
   Dim r_ado_Comand As ADODB.Command
   
   'Ejecuta los comandos SQL enviados por el cliente
   
   On Error GoTo Error_gf_EjecutaComandoSQL
    
   gf_EjecutaComandoSQL = 1
   Select Case Opcion
      Case 1
         'Ejecuta una consulta a la BD    (Por un Registro)
         Set p_RecordSet = New ADODB.Recordset
         p_RecordSet.Open p_Cadena, g_cnx_Conexion, adOpenStatic
         If p_RecordSet.EOF = False Then
             gf_EjecutaComandoSQL = 2
         End If
         Exit Function
      
      Case 2
         'Realiza la ejecución de un comando
         Set Command = New Command
         Command.ActiveConnection = g_cnx_Conexion
         Command.CommandText = "" & p_Cadena & ""
         Command.Execute
         Exit Function
       
      Case 3
         'Ejecuta un comando que puede o no devolver un conjunto de registros
         Set p_RecordSet = New ADODB.Recordset
         Set Command = New Command
         
         Command.ActiveConnection = g_cnx_Conexion
         Command.CommandText = "" & p_Cadena & ""
         
         'Set r_ado_Comand = New ADODB.Command
         
         
         'r_ado_Comand.CommandText = p_Cadena
         'r_ado_Comand.ActiveConnection = g_cnx_Conexion
         
         p_RecordSet.Open Command, , adOpenStatic, adLockReadOnly
         If p_RecordSet.State <> 1 Then
             gf_EjecutaComandoSQL = 0
         ElseIf p_RecordSet.EOF = False Then
             gf_EjecutaComandoSQL = 2
         End If
         Exit Function
   End Select

Error_gf_EjecutaComandoSQL:
   gf_EjecutaComandoSQL = -1
End Function


Public Function gf_CierraBaseDatos() As Boolean
   'Cierra la conexión con la Base de Datos
   
   On Error GoTo Error_gs_CierraBaseDatos
   
   g_cnx_Conexion.Close
   gf_CierraBaseDatos = True
   Exit Function

Error_gs_CierraBaseDatos:
   gf_CierraBaseDatos = False
   
End Function


