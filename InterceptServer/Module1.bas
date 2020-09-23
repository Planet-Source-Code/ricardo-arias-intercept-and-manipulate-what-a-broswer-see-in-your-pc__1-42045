Attribute VB_Name = "Module1"
Private Declare Function gethostbyname Lib "wsock32" _
  (ByVal hostname As String) As Long
  
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long

  
Public Function GetIPFromHostName(ByVal sHostName As String) As String


   Dim nbytes As Long
   Dim ptrHosent As Long
   Dim ptrName As Long
   Dim ptrAddress As Long
   Dim ptrIPAddress As Long
   Dim sAddress As String
   
   sAddress = Space$(4)

   ptrHosent = gethostbyname(sHostName & vbNullChar)

   If ptrHosent <> 0 Then

      ptrAddress = ptrHosent + 12
      
      CopyMemory ptrAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress, ByVal ptrAddress, 4
      CopyMemory ByVal sAddress, ByVal ptrIPAddress, 4

      GetIPFromHostName = IPToText(sAddress)

   End If
   
End Function

Private Function IPToText(ByVal IPAddress As String) As String

   IPToText = CStr(Asc(IPAddress)) & "." & _
              CStr(Asc(Mid$(IPAddress, 2, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 3, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 4, 1)))
              
End Function

Public Function FExist(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FExist = True
    Exit Function
MakeF:
        'error, file does Not exist
        FExist = False
    Exit Function
End Function
Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds


    Do Until GetTickCount > EndTime


        DoEvents
        Loop
End Function

