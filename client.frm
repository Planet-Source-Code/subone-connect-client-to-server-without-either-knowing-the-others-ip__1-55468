VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2700
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckBroadcast 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   0
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblRemotePort 
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1260
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblRemoteHostIP 
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1260
      TabIndex        =   6
      Top             =   660
      Width           =   1335
   End
   Begin VB.Label lblLocalPort 
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1260
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblLocalHostIP 
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1260
      TabIndex        =   4
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label lblStatic 
      Alignment       =   1  'Right Justify
      Caption         =   "RemotePort:"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblStatic 
      Alignment       =   1  'Right Justify
      Caption         =   "RemoteHostIP:"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label lblStatic 
      Alignment       =   1  'Right Justify
      Caption         =   "LocalPort:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblStatic 
      Alignment       =   1  'Right Justify
      Caption         =   "LocalHostIP:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1095
   End
End
Attribute VB_Name = "frmClient"
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
    
    'First we setup the client winsock
    With sckClient
        .Protocol = sckTCPProtocol
        .RemotePort = SERVER_PORT
    End With

    'Then the broadcast winsock
    With sckBroadcast
        .Protocol = sckUDPProtocol
        .LocalPort = CLIENT_BROADCAST_PORT
        .RemotePort = SERVER_BROADCAST_PORT
        .RemoteHost = "255.255.255.255"
        'This part is important, I'm not sure why, but if you
        'don't send a packet over the broadcast address from
        'one sock then it won't start receiving them from other
        'socks. I dunno if it's just VB, but if you know a
        'better way let me know. This is the part that some
        'other tutorials leave out BTW.
        .SendData ""
    End With
    
End Sub


Private Sub sckBroadcast_DataArrival(ByVal bytesTotal As Long)

    'Why would we want to create errors
    'over a packet that doesn't do anything?
    If bytesTotal = 0 Then Exit Sub
    
    Dim dat As String
    Dim param() As String
    
    'Extract the data to a local variable...
    sckBroadcast.GetData dat
    '...and seperate it into parameters
    param = Split(dat, DELIMITER)
    
    'Check that the message is from the server
    Select Case param(0)
        'That's him alright!
        Case IDENTIFY_SERVER

            'Check what the server wants
            Select Case CVL(param(1))
            
                'The server wants us to connect
                Case CONNECT_TO_HOST
                    If sckClient.State = sckClosed Then sckClient.Connect param(2)
                '}
                
            End Select
        
        '}
        
    End Select
    
End Sub


Private Sub sckClient_Connect()
    
    'Just confirm that we are really who he is looking for
    'by sending our signature and his IP back to him
    sckClient.SendData IDENTIFY_CLIENT & DELIMITER & _
                       MKL$(CONNECT_TO_HOST) & DELIMITER & _
                       sckClient.RemoteHostIP
                       
    'Show the connection info on our labels
    lblLocalHostIP.Caption = sckClient.LocalIP
    lblLocalPort.Caption = sckClient.LocalPort
    lblRemoteHostIP.Caption = sckClient.RemoteHostIP
    lblRemotePort.Caption = sckClient.RemotePort
    
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
