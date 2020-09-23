VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrBroadcast 
      Interval        =   1000
      Left            =   5520
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   5940
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckBroadcast 
      Left            =   6360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwConnections 
      Height          =   2295
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   4048
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblStatic 
      Caption         =   "&Connections:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Save some pain and suffering.
Option Explicit


'Set some port constants just in case we want to change
'the ports later without changing every one in code.
Private Const CLIENT_BROADCAST_PORT = 6125
Private Const SERVER_BROADCAST_PORT = 6126
Private Const SERVER_PORT = 6127

'Again, set some standards here to eliminate long hours
'of searching through code to change a constant value.
'(Oftentimes, simply using Find/Replace can do more
'damage than it does good)
Private Const DELIMITER = "åî"
Private Const CONNECT_TO_HOST = 1
Private Const IDENTIFY_SERVER = "BCs"
Private Const IDENTIFY_CLIENT = "BCc"


Private Sub Form_Load()

    If App.PrevInstance Then End

    'First we setup the ListView
    With lvwConnections
        .View = lvwReport
        .ColumnHeaders.Add , , "Socket", 1000
        .ColumnHeaders.Add , , "Port", 1000
        .ColumnHeaders.Add , , "IP Address", 2700
        .ColumnHeaders.Add , , "Confirmation", 2000
    End With
    
    'Then we setup the server winsock
    With sckServer(0)
        .Protocol = sckTCPProtocol
        .LocalPort = SERVER_PORT
        .Listen
    End With
    
    'Then the broadcast winsock
    With sckBroadcast
        .Protocol = sckUDPProtocol
        .LocalPort = SERVER_BROADCAST_PORT
        .RemotePort = CLIENT_BROADCAST_PORT
        .RemoteHost = "255.255.255.255"
    End With
    
End Sub


'Since we are sending a broadcast UDP packet, it will be
'received by us too, even though we are also the source.
'Unless we want these packets to pile up and waste space
'(or you have a use for this) then we need to clear them.
Private Sub sckBroadcast_DataArrival(ByVal bytesTotal As Long)
    
    Dim dat As String
    
    sckBroadcast.GetData dat
    
End Sub


Private Sub sckServer_Close(Index As Integer)

    lvwConnections.ListItems.Remove (lvwConnections.FindItem(Index, 0).Index)
    Unload sckServer(Index)

End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    
    Dim ub As Integer
    Dim lvi As ListItem
    
    'Load a new element in the array of sockets
    Load sckServer(sckServer.UBound + 1)
    ub = sckServer.UBound
    
    'Accept the connection on it
    sckServer(ub).Accept requestID
    
    'Add it to the list
    Set lvi = lvwConnections.ListItems.Add(, , ub)
    lvi.ListSubItems.Add , , sckServer(ub).LocalPort
    lvi.ListSubItems.Add , , sckServer(ub).RemoteHostIP
    lvi.ListSubItems.Add , , "Unconfirmed"
            
End Sub


Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    'Why would we want to create errors
    'over a packet that doesn't do anything?
    If bytesTotal = 0 Then Exit Sub
    
    Dim dat As String
    Dim param() As String
    Dim lvi As ListItem
    
    'Extract the data to a local variable...
    sckServer(Index).GetData dat
    '...and seperate it into parameters
    param = Split(dat, DELIMITER)
    
    'Check that the message is from the client
    Select Case param(0)
        'That's him alright!
        Case IDENTIFY_CLIENT

            'Check what the client wants
            Select Case CVL(param(1))
            
                'The client wants to confirm he is who we hope he is
                Case CONNECT_TO_HOST
                    Set lvi = lvwConnections.FindItem(Index, 0)
                    If sckServer(Index).LocalIP = param(2) Then lvi.ListSubItems(3).Text = "Passed" Else lvi.ListSubItems(3).Text = "Failed"
                '}
                
            End Select
        
        '}
        
    End Select
    
End Sub

'Do some stuff every second
Private Sub tmrBroadcast_Timer()

    'The first parameter of the protocol identifies the source
    'of the packet whether it be client (BCc) or server (BCs).
    'The second parameter is a command. In this case, simply
    'sending the identify string notifies the client that we
    'are anouncing the server's IP or hostname. Any further
    'parameters are specific to each command. In this case,
    'it simply contains the IP of the server to allow the
    'receiving client to connect to the server via TCP.
    
    'Send a broadcast packet to be picked up by all clients.
    sckBroadcast.SendData IDENTIFY_SERVER & DELIMITER & _
                          MKL$(CONNECT_TO_HOST) & DELIMITER & _
                          sckServer(0).LocalHostName
                              
End Sub


'Used to convert a 4 byte long integer into a 4 byte string
Private Function MKL$(ByVal lngNum As Long)

    MKL$ = Chr$(lngNum And &HFF&) & _
           Chr$((lngNum And &HFF00&) / &H100&) & _
           Chr$((lngNum And &HFF0000) / &H10000) & _
           Chr$((lngNum And &HFF000000) / &H1000000)
           
End Function


'Used to convert a 4 byte string into a 4 byte long integer
Private Function CVL(ByVal strNum As String) As Long

    CVL = Asc(Mid$(strNum, 1, 1)) + _
          Asc(Mid$(strNum, 2, 1)) * &H100& + _
          Asc(Mid$(strNum, 3, 1)) * &H10000 + _
          Asc(Mid$(strNum, 4, 1)) * &H1000000
    
End Function
