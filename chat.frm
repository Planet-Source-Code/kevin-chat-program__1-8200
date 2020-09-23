VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chat Program"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&About"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "................."
      Top             =   480
      Width           =   5295
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "................."
      Top             =   120
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF8080&
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "chat.frx":0000
      Top             =   3360
      Width           =   7095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF8080&
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "chat.frx":0043
      Top             =   1080
      Width           =   7095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3240
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   5567
      LocalPort       =   5567
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "This Chat program is maked by: Kevin  (kevin_verp@hotpop.com)"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   4650
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your IP Address:"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chat with: (IP Address:)"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Text:"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Their Text:"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************'
'*  Chat Program maked by: Kevin                         *'
'*  E-Mail: kevin_verp@hotpop.com                        *'
'*  Comments:                                            *'
'*  ---------                                            *'
'*  Have Fun !                                           *'
'*********************************************************'
Private Sub Command1_Click()
Winsock1.SendData Text4.Text + " has left the chat !"
Winsock1.Close
End
End Sub
Private Sub Command2_Click()
MsgBox "If you want to use this in your own programs, please contact me !", vbInformation, "About"
If Label5.Top = 5520 Then
Label5.Top = 6120
Else
Label5.Top = 5520
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Text1.Text = "> Chat Program by: Kevin (kevin_verp@hotpop.com)"
Winsock1.RemoteHost = Winsock1.LocalIP
Winsock1.SendData ">"
Text4.Text = Winsock1.LocalIP
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Text1 = Text1
Winsock1.SendData (KeyAscii)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
Text2 = Text2
Winsock1.GetData (KeyAscii)
End Sub
Private Sub Text3_Change()
Winsock1.RemoteHost = Text3.Text
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Winsock1.GetData strData, vbString
Text3.Text = Winsock1.RemoteHostIP
Winsock1.RemoteHost = Winsock1.RemoteHostIP
If Asc(strData) = 8 And Len(Text2) > 0 Then
Text2.Text = Mid(Text2, 1, (Len(Text2) - 1))
Else
Text2 = Text2 & strData
End If
If Asc(strData) = 13 Then
Text2 = Text2 & vbNewLine
End If
Text2.SelStart = Len(Text2)
End Sub

