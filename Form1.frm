VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Whois by BugMaster"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "lookup"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFC0&
      Height          =   355
      Left            =   1560
      TabIndex        =   0
      Text            =   "turk.net"
      Top             =   120
      Width           =   3495
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5400
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtResponse 
      BackColor       =   &H00C0C0FF&
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   600
      Width           =   7935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter search string:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   200
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    txtSearch = ""
    txtResponse = ""
End Sub

Private Sub Command4_Click()
   MousePointer = vbHourglass
   txtResponse = ""
   Winsock1.Close
   Winsock1.LocalPort = 0
   If Right(txtSearch, 3) = ".tr" Then
      Winsock1.Connect "whois.metu.edu.tr", 43
   Else
      Winsock1.Connect "rs.internic.net", 43
   End If
   
End Sub

Private Sub Winsock1_Connect()
    Winsock1.SendData txtSearch & vbCrLf
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String

    On Error Resume Next

    Winsock1.GetData strData
    strData = Replace(strData, Chr$(10), vbCrLf)
    txtResponse = txtResponse & strData
    MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub
