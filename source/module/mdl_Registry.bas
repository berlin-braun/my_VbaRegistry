Attribute VB_Name = "mdl_Registry"
Option Compare Database
Option Explicit


Public Function registry_Key_Read(ByVal str_Key As String) As String
  Dim str_Ret   As String
  Dim ws        As Object
  
On Error GoTo my_Error

  str_Ret = ""
  
  Set ws = CreateObject("WScript.Shell")
  str_Ret = ws.RegRead(str_Key)


my_Exit:
    
  On Error Resume Next
  
  Set ws = Nothing
  registry_Key_Read = str_Ret
  
  On Error GoTo 0
  Exit Function

my_Error:
  Dim str_Error  As String
  Dim lng_Error  As Long

  lng_Error = Err.Number
  str_Error = "Error " & Err.Number & " (" & Err.Description & ") in procedure registry_Key_Read of Modul mod_Registry"

  Select Case lng_Error
    Case Is = 0: Resume Next

    Case Else:  MsgBox str_Error
  End Select
  GoTo my_Exit

End Function


Public Function registry_Key_Exists(ByVal str_Key As String) As Boolean
  Dim bln_Ret   As Boolean
  Dim str_Tmp   As String
  Dim ws        As Object
  
On Error GoTo my_Error

  bln_Ret = False
  str_Tmp = ""
  
  Set ws = CreateObject("WScript.Shell")
  str_Tmp = ws.RegRead(str_Key)
  
  If Len(str_Tmp) > 0 Then
    bln_Ret = True
  End If

my_Exit:
  
  registry_Key_Exists = bln_Ret
  
  On Error GoTo 0
  Exit Function

my_Error:
  Dim str_Error  As String
  Dim lng_Error  As Long


  Select Case lng_Error
    Case Is = 0: Resume Next
    Case Is = -2147024894 ': Resume Next            ' Key nor found
    Case Else:
      MsgBox str_Error
  
  End Select
  GoTo my_Exit

End Function

