VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debug"
   ClientHeight    =   1935
   ClientLeft      =   8310
   ClientTop       =   2160
   ClientWidth     =   3810
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3810
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock WinS 
      Left            =   2520
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   2710
      LocalPort       =   2710
   End
   Begin VB.Frame Frame2 
      Caption         =   "Your turn"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "y or n"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data to tell what to put on which button"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.Label Label2 
         Caption         =   "Which caption"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "X or O"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide 'Hides the form (DUR)
End Sub

Private Sub Form_Load()
    MakeTransparent Me.hWnd, 150 'Makes the form transparent
End Sub

Private Sub WinS_Close()
Form1.Label1.Caption = "Connection closed" 'If the connection closes it tells
 'you on the status bar
End Sub

Private Sub WinS_Connect() 'When it connects it tells you and enables
 'the buttons
Form1.Label1.Caption = "Connected...  Opponent's turn"
Form1.CMD1.Enabled = True
Form1.CMD2.Enabled = True
Form1.CMD3.Enabled = True
Form1.CMD4.Enabled = True
Form1.CMD5.Enabled = True
Form1.CMD6.Enabled = True
Form1.CMD7.Enabled = True
Form1.CMD8.Enabled = True
Form1.CMD9.Enabled = True
End Sub

Private Sub WinS_ConnectionRequest(ByVal requestID As Long) 'When your
WinS.Close 'opponent connects to you it accepts the request, tells you
WinS.Accept requestID 'its your turn, and enables all the buttons
Form1.CMD1.Enabled = True
Form1.CMD2.Enabled = True
Form1.CMD3.Enabled = True
Form1.CMD4.Enabled = True
Form1.CMD5.Enabled = True
Form1.CMD6.Enabled = True
Form1.CMD7.Enabled = True
Form1.CMD8.Enabled = True
Form1.CMD9.Enabled = True
Form1.Label1.Caption = "Connected... Your turn"
End Sub

Private Sub WinS_DataArrival(ByVal bytesTotal As Long) 'Lots of stuff here
Dim inComing As String 'Dims things
WinS.GetData inComing 'Splits the message so it can see what to do. Later
Label1.Caption = Split(inComing, "+")(0) 'I'll add a chat and label1 will be
Label2.Caption = Split(inComing, "+")(1) 'chat instead of X
checkFORwiNNer 'Checks to see if someone has a winning or loosing board
Text1.Text = "y" 'Makes it your turn
Form1.Label1.Caption = "Recieved opponent's move...  Your turn" 'Tells you its
If Label1.Caption = "X" And Label2.Caption = "1" Then           'your turn
Form1.CMD1.Caption = "O" 'Checks to se where to put the O
ElseIf Label1.Caption = "X" And Label2.Caption = "2" Then
Form1.CMD2.Caption = "O"
ElseIf Label1.Caption = "X" And Label2.Caption = "3" Then
Form1.CMD3.Caption = "O"
ElseIf Label1.Caption = "X" And Label2.Caption = "4" Then
Form1.CMD4.Caption = "O"
ElseIf Label1.Caption = "X" And Label2.Caption = "5" Then
Form1.CMD5.Caption = "O"
ElseIf Label1.Caption = "X" And Label2.Caption = "6" Then
Form1.CMD6.Caption = "O"
ElseIf Label1.Caption = "X" And Label2.Caption = "7" Then
Form1.CMD7.Caption = "O"
ElseIf Label1.Caption = "X" And Label2.Caption = "8" Then
Form1.CMD8.Caption = "O"
ElseIf Label1.Caption = "X" And Label2.Caption = "9" Then
Form1.CMD9.Caption = "O"
End If

If Label1.Caption = "a" And Label2.Caption = "ss" Then 'If your opponent wants
Dim yesORno As VbMsgBoxResult                          'to restart it asks you
yesORno = MsgBox("Your opponent wants to restart the game.  Do you agree?", vbYesNo, "Restart?")
If yesORno = vbYes Then 'If you agree the game restarts
Form1.CMD1.Caption = ""
Form1.CMD2.Caption = ""
Form1.CMD3.Caption = ""
Form1.CMD4.Caption = ""
Form1.CMD5.Caption = ""
Form1.CMD6.Caption = ""
Form1.CMD7.Caption = ""
Form1.CMD8.Caption = ""
Form1.CMD9.Caption = ""
WinS.SendData "ok+dumbass" 'Sends a confermation to your opponent
Form3.Text1.Text = "y" 'Makes it your turn
End If

If Label1.Caption = "ok" And Label2.Caption = "dumbass" Then 'If your
MsgBox "Your opponent has agreed to restart the game", , "Restart" 'opponent
Form1.CMD1.Caption = ""                                        'aggres, the
Form1.CMD2.Caption = ""                                       'game restarts
Form1.CMD3.Caption = ""
Form1.CMD4.Caption = ""
Form1.CMD5.Caption = ""
Form1.CMD6.Caption = ""
Form1.CMD7.Caption = ""
Form1.CMD8.Caption = ""
Form1.CMD9.Caption = ""
Form3.Text1.Text = "n" 'Makes it not your turn
End If
End If

checkFORwiNNer 'Checks to see if someone has a winning or loosing board, again

End Sub

Private Sub WinS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Form1.Label1.Caption = "There was an error" 'Tells you if there was an error
 'connecting
End Sub

Private Sub WinS_SendComplete()
checkFORwiNNer 'Checks to see if someone has a winning or loosing board
Form1.Label1.Caption = "Sent move...  Opponent's turn" 'Tells you the move was
 'sent
End Sub

Private Sub WinS_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
checkFORwiNNer 'Checks to see if someone has a winning or loosing board
Form1.Label1.Caption = "Sending move..." 'Unless your computer is slow, you'll
'never see this.  Winsock is to fast and simple
End Sub
