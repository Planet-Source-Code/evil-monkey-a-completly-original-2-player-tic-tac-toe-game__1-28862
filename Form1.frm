VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic-Tac Toe"
   ClientHeight    =   3150
   ClientLeft      =   4740
   ClientTop       =   3720
   ClientWidth     =   3225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3225
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   2400
   End
   Begin VB.CommandButton CMD9 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton CMD8 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton CMD7 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton CMD6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CMD5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CMD4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CMD3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton CMD2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton CMD1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   375
      Left            =   80
      TabIndex        =   0
      Top             =   2925
      Width           =   4575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuListen 
         Caption         =   "&Listen"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestart 
         Caption         =   "&Restart game"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "&Debug"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'All this stuff is for making the hidden manifest file
 'so it looks like Windows XP.  If you dont have XP, you might as well delete
 'this stuff
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long


Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
    Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Sub CMD1_Click()
On Error GoTo Hell 'Error handeling
If Form3.Text1.Text = "" Then
GoTo Hell 'If the game hasn't begun it goes to "Hell" (More error and cheating
End If 'handeling)
If Form3.Text1.Text = "y" Then 'Checks to see if its their turn
If CMD1.Caption = "X" Then 'If theres an X there it says you cant go there.
MsgBox "You can't go there", , "Error"
Exit Sub
End If
If CMD1.Caption = "O" Then 'If theres an O there it says you cant go there.
MsgBox "You can't go there", , "Error"
Exit Sub
End If
CMD1.Caption = "X" 'Puts an X
Form3.WinS.SendData "X+1" 'and sends a message to put an O on the other
Form3.Text1.Text = "n" 'person's game
Else
Beep 'Beeps and says its not their turn
Label1.Caption = "It's not your turn"
End If
Exit Sub
Hell: 'Error handeling
MsgBox "You are not connected to a remote player", vbCritical, "Error"
End Sub

Private Sub CMD2_Click() 'Read the comments for CMD1, please
On Error GoTo Hell
If Form3.Text1.Text = "" Then
GoTo Hell
End If
If Form3.Text1.Text = "y" Then
If CMD2.Caption = "X" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
If CMD2.Caption = "O" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
CMD2.Caption = "X"
Form3.WinS.SendData "X+2"
Form3.Text1.Text = "n"
Else
Beep
Label1.Caption = "It's not your turn"
End If
Exit Sub
Hell:
MsgBox "You are not connected to a remote player", vbCritical, "Error"
End Sub

Private Sub CMD3_Click() 'Read the comments for CMD1, please
On Error GoTo Hell
If Form3.Text1.Text = "" Then
GoTo Hell
End If
If Form3.Text1.Text = "y" Then
If CMD3.Caption = "X" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
If CMD3.Caption = "O" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
CMD3.Caption = "X"
Form3.WinS.SendData "X+3"
Form3.Text1.Text = "n"
Else
Beep
Label1.Caption = "It's not your turn"
End If
Exit Sub
Hell:
MsgBox "You are not connected to a remote player", vbCritical, "Error"
End Sub

Private Sub CMD4_Click() 'Read the comments for CMD1, please
On Error GoTo Hell
If Form3.Text1.Text = "" Then
GoTo Hell
End If
If Form3.Text1.Text = "y" Then
If CMD4.Caption = "X" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
If CMD4.Caption = "O" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
CMD4.Caption = "X"
Form3.WinS.SendData "X+4"
Form3.Text1.Text = "n"
Else
Beep
Label1.Caption = "It's not your turn"
End If
Exit Sub
Hell:
MsgBox "You are not connected to a remote player", vbCritical, "Error"
End Sub

Private Sub CMD5_Click() 'Read the comments for CMD1, please
On Error GoTo Hell
If Form3.Text1.Text = "" Then
GoTo Hell
End If
If Form3.Text1.Text = "y" Then
If CMD5.Caption = "X" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
If CMD5.Caption = "O" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
CMD5.Caption = "X"
Form3.WinS.SendData "X+5"
Form3.Text1.Text = "n"
Else
Beep
Label1.Caption = "It's not your turn"
End If
Exit Sub
Hell:
MsgBox "You are not connected to a remote player", vbCritical, "Error"
End Sub

Private Sub CMD6_Click() 'Read the comments for CMD1, please
On Error GoTo Hell
If Form3.Text1.Text = "" Then
GoTo Hell
End If
If Form3.Text1.Text = "y" Then
If CMD6.Caption = "X" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
If CMD6.Caption = "O" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
CMD6.Caption = "X"
Form3.WinS.SendData "X+6"
Form3.Text1.Text = "n"
Else
Beep
Label1.Caption = "It's not your turn"
End If
Exit Sub
Hell:
MsgBox "You are not connected to a remote player", vbCritical, "Error"
End Sub

Private Sub CMD7_Click() 'Read the comments for CMD1, please
On Error GoTo Hell
If Form3.Text1.Text = "" Then
GoTo Hell
End If
If Form3.Text1.Text = "y" Then
If CMD7.Caption = "X" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
If CMD7.Caption = "O" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
CMD7.Caption = "X"
Form3.WinS.SendData "X+7"
Form3.Text1.Text = "n"
Else
Beep
Label1.Caption = "It's not your turn"
End If
Exit Sub
Hell:
MsgBox "You are not connected to a remote player", vbCritical, "Error"
End Sub

Private Sub CMD8_Click() 'Read the comments for CMD1, please
On Error GoTo Hell
If Form3.Text1.Text = "" Then
GoTo Hell
End If
If Form3.Text1.Text = "y" Then
If CMD8.Caption = "X" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
If CMD8.Caption = "O" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
CMD8.Caption = "X"
Form3.WinS.SendData "X+8"
Form3.Text1.Text = "n"
Else
Beep
Label1.Caption = "It's not your turn"
End If
Exit Sub
Hell:
MsgBox "You are not connected to a remote player", vbCritical, "Error"
End Sub

Private Sub CMD9_Click() 'Read the comments for CMD1, please
On Error GoTo Hell
If Form3.Text1.Text = "" Then
GoTo Hell
End If
If Form3.Text1.Text = "y" Then
If CMD9.Caption = "X" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
If CMD9.Caption = "O" Then
MsgBox "You can't go there", , "Error"
Exit Sub
End If
CMD9.Caption = "X"
Form3.WinS.SendData "X+9"
Form3.Text1.Text = "n"
Else
Beep
Label1.Caption = "It's not your turn"
End If
Exit Sub
Hell:
MsgBox "You are not connected to a remote player", vbCritical, "Error"
End Sub

Private Sub Form_Initialize() 'Loads the Windows XP manifest file
    Dim xptheme As Long 'so all the buttons and stuff are XP themed
    Dim manifestpth As String
On Error GoTo manifestdoesnotexisT 'If theres no manifest file it will error,
                                   'this will send it to the error handeling
                                   'that will make one

    If Right(App.Path, 1) = "\" Then 'Loads the file
        manifestpth = App.Path & App.EXEName & ".exe.manifest"
    Else
        manifestpth = App.Path & "\" & App.EXEName & ".exe.manifest"
    End If
    FileCopy manifestpth, "c:\checkexist.txt"
    Kill "c:\checkexist.txt"
    xptheme = InitCommonControls
    Exit Sub
manifestdoesnotexisT:
    Call makeNEWmanifest 'Makes a new manifest file
End Sub
Sub makeNEWmanifest() 'Makes a new manifest file
    Dim NEWmanifestpth As String
    Dim xptheme As Long
    Dim setAShidden As Long
    On Error GoTo problemARGH 'ARGH!


    If Right(App.Path, 1) = "\" Then 'Sets the file path for the file
        NEWmanifestpth = App.Path & App.EXEName & ".exe.manifest"
    Else
        NEWmanifestpth = App.Path & "\" & App.EXEName & ".exe.manifest"
    End If
    Open NEWmanifestpth For Output As #1 'Makes it
    Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?><assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & "><assemblyIdentity version=" & Chr(34) & "1.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " name=" & Chr(34) & "HybridDesign.WindowsXP.Example" & Chr(34) & " type=" & Chr(34) & "win32" & Chr(34) & " /> <description>An example of windows XP theming.</description> <dependency> <dependentAssembly> <assemblyIdentity type=" & Chr(34) & "win32" & Chr(34) & " name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34) & " version=" & Chr(34) & "6.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34) & " language=" & Chr(34) & "*" & Chr(34) & " /> </dependentAssembly> </dependency> </assembly>"
    Close #1
    xptheme = InitCommonControls 'Makes the file hidden
    setAShidden = SetFileAttributes(NEWmanifestpth, FILE_ATTRIBUTE_HIDDEN)
    Timer1.Enabled = True 'Enables the timer that will restart the program
    Exit Sub
problemARGH: 'Telling the person that themes will not be enabled
    MsgBox "Error creating Windows XP theme file. You may be running EXE file from a network drive with which you dont have write permissions. Themes will not be enabled.", vbExclamation, "Themeing Error!"
End Sub

Private Sub Form_Load()
MakeTransparent Me.hWnd, 150 'Makes the form transparent (Cool)
End Sub

Private Sub mnuConnect_Click()
Form2.Show 'Shows the server form and selects the connect option
Form2.Option2.Value = True
End Sub

Private Sub mnuDebug_Click()
Form3.Show 'Shows the debug form
End Sub

Private Sub mnuExit_Click()
End 'Exits the program
End Sub

Private Sub mnuListen_Click()
Form2.Show 'Shows the server form and selects the listen option
Form2.Option1.Value = True
End Sub

Private Sub mnuRestart_Click()
On Error GoTo Hell 'Error handeling
Form3.WinS.SendData "a+ss" 'Ass
Exit Sub
Hell:
MsgBox "You aren't connected to a remote opponent", vbCritical, "Error"
End Sub

Private Sub Timer1_Timer() 'Restarts the program so people can see the XP theme
    On Error GoTo Error 'Error handeling
    Dim myEXEpath As String
    
    Unload Form1


    If Right(App.Path, 1) = "\" Then
        myEXEpath = App.Path & App.EXEName & ".exe"
    Else
        myEXEpath = App.Path & "\" & App.EXEName & ".exe"
    End If
    Shell myEXEpath, vbNormalFocus
    Exit Sub
    Timer1.Enabled = False 'Disables the timer so it doesnt go on a loop (That
                           'would be suc)
Error: 'You can't run the program in VB
    MsgBox "Error exucuting the EXE file. This would be caused by you trying to compile the manifest file from inside Visual Basic. You can only see the theme when fully compiled, and ran as an .EXE file :)", vbExclamation, "Manifest Exucution Error!"
End Sub
