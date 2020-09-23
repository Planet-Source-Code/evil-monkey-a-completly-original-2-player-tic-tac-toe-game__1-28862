VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   2895
   ClientLeft      =   6510
   ClientTop       =   5190
   ClientWidth     =   2775
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   2775
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "Remote Port"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "Remote IP"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Connect"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Listen"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Local Port"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2400
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connect"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listen"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2535
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Your IP : 00.00.00.00"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'Decides what to do, listen or connect, then
                             'does it
Form2.Hide 'Hides the server form so it doesn't stay there and be gay
On Error GoTo Hell1 'Error handeling number one
If Option1.Value = True Then 'If listen is selected
Form3.WinS.Close
Form3.WinS.LocalPort = Text1.Text 'Sets the port
Form3.WinS.Listen 'Listens
Form3.Text1.Text = "y" 'Makes it your turn
Form1.Label1.Caption = "Listening..." 'Changes the status bar
Else
On Error GoTo Hell2 'Error handeling number two
Form3.WinS.Close
Form3.WinS.Connect Text2.Text, Text3.Text 'Connects to the IP and the port
Form3.Text1.Text = "n" 'Makes it not your turn
Form1.Label1.Caption = "Connecting..." 'Changes the status bar
End If
Exit Sub
Hell1: 'Error handeling number one
MsgBox "You need to enter a port number  I.E. 1234", vbCritical, "Error"
Form3.Text1.Text = ""
Exit Sub
Hell2: 'Error handeling number two
MsgBox "Could not connect", vbCritical, "Error"
Form3.Text1.Text = ""
End Sub

Private Sub Form_Load() 'Makes the form transparent and shows what your IP is
MakeTransparent Me.hWnd, 150
Label1.Caption = "Your IP : " & Form3.WinS.LocalIP
End Sub
