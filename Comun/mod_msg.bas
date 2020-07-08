Attribute VB_Name = "modmsg"
   Option Explicit
   
   ' Declaraciones del Api
   '******************************************************************************
   ' Establece el Hook
   Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
   ByVal idHook As Long, _
   ByVal lpfn As Long, _
   ByVal hmod As Long, _
   ByVal dwThreadId As Long) As Long
   
   ' Destruye el Hook
   Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
   ByVal hHook As Long) As Long
   
   ' Cambia el texto al bot�n del Msgbox
   Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" ( _
   ByVal hDlg As Long, _
   ByVal nIDDlgItem As Long, _
   ByVal lpString As String) As Long
   
   
   'Contantes
   Private Const WH_CBT = 5
   Private Const HCBT_ACTIVATE = 5
   
   'Enumeraciones para el bot�n que se va a modificar
   Enum Ebuttons
      ' Para el Bot�n OK
      [OK] = 1
      ' Para el Bot�n Cancelar
      [Cancel] = 2
      ' Para el Bot�n Abortar
      [ABORT] = 3
      ' Para el Bot�n Reintentar
      [RETRY] = 4
      ' Para el Bot�n Ignorar
      [Ignore] = 5
      ' Para el Bot�n Si
      [YES] = 6
      ' Para el Bot�n No
      [NO] = 7
   End Enum
   
   ' variables que toman los valores de la funci�n _
   MsgBoxExText y los usa dentro del HOOK
   
   Dim r_cmd_Boton1 As Long ' Elbot�n que se va a modificar
   Dim r_str_TexBo1 As String ' Texto del bot�n
   
   Dim r_cmd_Boton2 As Long ' Elbot�n que se va a modificar
   Dim r_str_TexBo2 As String ' Texto del bot�n
   
   Dim r_cmd_Boton3 As Long ' Elbot�n que se va a modificar
   Dim r_str_TexBo3 As String ' Texto del bot�n
   
   
   ' Mantiene el valor para luego finalizar el Hook
   Private Id_Hook As Long
   
   Function MsgBoxExText(Prompt As String, Buttons As VbMsgBoxStyle, Title As String, _
      p_Boton1 As Ebuttons, _
      p_TexBo1 As String, _
      Optional p_Boton2 As Ebuttons, _
      Optional p_TexBo2 As String, _
      Optional p_Boton3 As Ebuttons, _
      Optional p_TexBo3 As String) As VbMsgBoxResult
      
      r_cmd_Boton1 = p_Boton1
      r_str_TexBo1 = p_TexBo1
      
      r_cmd_Boton2 = p_Boton2
      r_str_TexBo2 = p_TexBo2
      
      r_cmd_Boton3 = p_Boton3
      r_str_TexBo3 = p_TexBo3
      
      Hook
      MsgBoxExText = MsgBox(Prompt, Buttons, Title)
      
   End Function
   
   Private Sub Hook()
      
      ' Inicia el Hook
      Id_Hook = SetWindowsHookEx(WH_CBT, AddressOf winProc, 0, App.ThreadID)
      
   End Sub
   
   ' Procedimiento que intercepta los mensajes
   Public Function winProc( _
      ByVal uMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long) As Long
      
      Dim ret As Long
      
      If uMsg = HCBT_ACTIVATE Then
      
      ' Cambia el texto
      ret = SetDlgItemText(wParam, r_cmd_Boton1, r_str_TexBo1)
      
      If r_str_TexBo2 <> "" Then
         ret = SetDlgItemText(wParam, r_cmd_Boton2, r_str_TexBo2)
      End If
      
      If r_str_TexBo3 <> "" Then
         ret = SetDlgItemText(wParam, r_cmd_Boton3, r_str_TexBo3)
      End If
      
      ' Elimina el Hook
      ret = UnhookWindowsHookEx(Id_Hook)
      r_str_TexBo1 = vbNullString
      
      End If
      
      winProc = 0
   
   End Function
   
   
