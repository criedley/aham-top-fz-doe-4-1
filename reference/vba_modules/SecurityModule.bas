Attribute VB_Name = "SecurityModule"
    ' Makes sure all variables are dimensioned in each subroutine.
     Option Explicit

     ' Access the GetUserNameA function in advapi32.dll and
     ' call the function GetUserName.
     Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
     (ByVal lpBuffer As String, nSize As Long) As Long
     
     Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias _
        "GetComputerNameA" (ByVal lpBuffer As String, _
        nSize As Long) As Long
        
Sub Get_Computer_Name()
   Dim strBuffer As String * 255
   If GetComputerName(strBuffer, 255&) <> 0 Then
      ' Name exist
      ComputerName = Left(strBuffer, InStr(strBuffer, vbNullChar) - 1)
   Else
      ' Name doesn't exist
      ComputerName = "N/A"
   End If
End Sub

' Main routine to Dimension variables, retrieve user name
' and display answer.
Sub Get_User_Name()
    ' Dimension variables
    Dim lpBuff As String * 25
    ' Get the user name minus any trailing spaces found in the name.
    'ret = GetUserName(lpBuff, 25)
    If GetUserName(lpBuff, 25) <> 0 Then
        UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    Else
        UserName = "N/A"
    End If
    ' Display the User Name
End Sub

