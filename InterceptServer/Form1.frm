VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
   Begin MSWinsockLib.Winsock Intermediary 
      Index           =   0
      Left            =   5760
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Intercept 
      Index           =   0
      Left            =   5280
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Ws 
      Left            =   4800
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim IP(100) As String
Public Header As String
Dim nIP As String
Private Sub Form_Load()
On Error Resume Next
'Find Host File
Dim nHost As String
If FExist("c:\windows\system32\drivers\etc\hosts") = True Then
    nHost = "c:\windows\system32\drivers\etc\"
ElseIf FExist("c:\windows\hosts") = True Then
    nHost = "c:\windows\"
ElseIf FExist("c:\winnt\system32\drivers\etc\hosts") = True Then
    nHost = "c:\winnt\system32\drivers\etc\"
End If

'Read the sites that you want to intercept
Dim nSites As String
Open App.Path + "\Sites.txt" For Binary Access Read As #1
nSites = Space(LOF(1))
Get #1, , nSites
Close #1

nSites = nSites + Chr(13) + Chr(10) 'if the last line dont have it!


'Split it
Dim aSites() As String, xSites() As String
aSites = Split(nSites, Chr(13) + Chr(10))
xSites = Split(nSites, Chr(13) + Chr(10))
'Check if the host file already has your sites and erase it in  order to get the IP
Dim eHost As String
Open nHost & "hosts" For Binary Access Read As #1
eHost = Space(LOF(1))
Get #1, , eHost
Close #1

For i = 0 To UBound(aSites)
    If aSites(i) <> "" Then
        aSites(i) = "127.0.0.1 " + aSites(i)
        List1.AddItem aSites(i), i
    End If
Next i

eHost = Replace(eHost, "# Sites that will be intercepted" + Chr(13) + Chr(10), "")

For i = 0 To UBound(aSites)
    If aSites(i) <> "" Then
        eHost = Replace(LCase(eHost), LCase(aSites(i)) + Chr(13) + Chr(10), "")
    End If
Next i


'To get the IPs need a host file with out the sites that will be intercepted
Kill nHost & "hosts"

Open nHost & "hosts" For Binary Access Write As #2
Put #2, , eHost
Close #2

Wait 1 '

For i = 0 To UBound(aSites)
    If aSites(i) <> "" Then
        IP(i) = GetIPFromHostName(xSites(i))
    End If
Next i


'Write to the host file the sites that you want to intercept
Open nHost & "hosts" For Append As #1 'OPEN the HOSTS File For Appending Write
    Print #1, vbCrLf 'Print a Carrige Return
    Print #1, "# Sites that will be intercepted"
    
    Do Until X = List1.ListCount 'Begin a Do Until Loop
    DoEvents 'Do events
        Print #1, List1.List(X) 'Append the Server from the Listbox
        X = X + 1 'add 1 to X so next leep we will get onto the NEXT ITEM in the List
    Loop 'loop (for the new to VB people, it goes back to "Print #1, List1.List(X)" 3 lines up
    Close #1 'Close the file, save the writing

Ws.LocalPort = 80 ' The port for normal web browser
Ws.Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Live host file as where at the beginning, with out our changes : )
On Error Resume Next

Dim nHost As String
If FExist("c:\windows\system32\drivers\etc\hosts") = True Then
    nHost = "c:\windows\system32\drivers\etc\"
ElseIf FExist("c:\windows\hosts") = True Then
    nHost = "c:\windows\"
ElseIf FExist("c:\winnt\system32\drivers\etc\hosts") = True Then
    nHost = "c:\winnt\system32\drivers\etc\"
End If

Dim nSites As String
Open App.Path + "\Sites.txt" For Binary Access Read As #1
nSites = Space(LOF(1))
Get #1, , nSites
Close #1

nSites = nSites + Chr(13) + Chr(10) 'if the last line dont have it!

Dim aSites() As String, xSites() As String
aSites = Split(nSites, Chr(13) + Chr(10))
xSites = Split(nSites, Chr(13) + Chr(10))
'Check if the host file already has your sites and erase it in  order to get the IP
Dim eHost As String
Open nHost & "hosts" For Binary Access Read As #1
eHost = Space(LOF(1))
Get #1, , eHost
Close #1

For i = 0 To UBound(aSites)
    If aSites(i) <> "" Then
        aSites(i) = "127.0.0.1 " + aSites(i)
    End If
Next i
'Clean it
eHost = Replace(eHost, "# Sites that will be intercepted" + Chr(13) + Chr(10), "")

For i = 0 To UBound(aSites)
    If aSites(i) <> "" Then
        eHost = Replace(LCase(eHost), LCase(aSites(i)) + Chr(13) + Chr(10), "")
    End If
Next i


Kill nHost & "hosts"
'Save new CLEAN file
Open nHost & "hosts" For Binary Access Write As #2
Put #2, , eHost
Close #2
For i = 0 To 100
Unload Intercept(i)
Unload Intermediary(i)
Next i
End Sub

Private Sub Intercept_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim Dat As String
Dat = Space(bytesTotal)
Intercept(Index).GetData Dat, , bytesTotal 'Receive message
If GetHost(Dat) <> 0 Then
    If Intermediary(Index).State = 0 Then 'some times fails saying that the control index dont exists
        Intermediary(Index).Connect nIP, 80
    Else
        Intermediary(Index).SendData Header
    End If
End If
End Sub

Public Function GetHost(Headers As String) As Long
Dim Pos As Long, Pos1 As Long, Result As String
Pos = InStr(LCase(Headers), "host:")
Pos1 = InStr(Pos, Headers, Chr(13))
Result = Mid$(Headers, Pos + 6, (Pos1 - 1) - (Pos + 5))
For i = 0 To List1.ListCount
    If InStr(LCase(List1.List(i)), LCase(Result)) Then
    Header = Replace(Headers, Result, IP(i))
    Me.Caption = IP(i)
    nIP = IP(i)
    GetHost = 1
    End If
    DoEvents
Next i
End Function

Private Sub Intermediary_Connect(Index As Integer)
Me.Caption = "Connected"
Intermediary(Index).SendData Header
End Sub

Private Sub Intermediary_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Me.Caption = "DataArrival"
Dim Dat1 As String, Estado As Long
Dat1 = Space(bytesTotal)
Intermediary(Index).GetData Dat1, , bytesTotal 'Receive message

'Here is where you can HACK the page doing replace
'in a near future it will be handled more easily.
'now do it manually

'examples
'Google:
Dat1 = Replace(Dat1, "Feeling Lucky", "Feel Hacked!!") ' Change the original text received from google
Dat1 = Replace(Dat1, "</table></form><br>", "</table></form><br><center><font color='#ff0000'><h1>Google Hacked!!!</h1></font><p><a href='http:\\www.planetsourcecode.com'> Dont forget to vote for my code!!!</a></center>")

'Other pages
'Dat1 = Replace(LCase(Dat1), "<body>", "<body><p><center><b><h1>Hacked!!</h1></b></center><p>")
'Dat1 = Replace(LCase(Dat1), "</body>", "<p><center><b><h1>Hacked!!</h1></b></center><p></body>")
'Dat1 = Replace(LCase(Dat1), "</center>", "</center><p><center><b><h1>Hacked!!</h1></b></center><p>")

Otra:
Estado = Intercept(Index).State
DoEvents
Intercept(Index).SendData Dat1
Me.Caption = "Enviando"
If Estado = 9 Then GoTo Otra
Me.Caption = "Salio de otra"
End Sub

Private Sub Ws_ConnectionRequest(ByVal requestID As Long)
Dim Server As Control, Intermed As Control
Set Server = Intercept(Intercept.UBound + 1)
DoEvents
Set Intermed = Intermediary(Intermediary.UBound + 1)
DoEvents
Load Server
DoEvents
Load Intermed
DoEvents
Server.Accept requestID 'accept any request
End Sub



