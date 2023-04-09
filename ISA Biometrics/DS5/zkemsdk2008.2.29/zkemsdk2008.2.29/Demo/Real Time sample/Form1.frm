VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   615
      Left            =   150
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOpen_Click()
    'If CZKEM1.Connect_Com(1, 1, 115200) Then
    If CZKEM1.Connect_Net("192.168.1.251", 4370) Then
        cmdClose.Enabled = True
        cmdOpen.Enabled = False
    Else
        Text1.Text = Text1.Text & "Connect fail" & Chr(13) & Chr(10)
    End If
End Sub

Private Sub cmdClose_Click()
    CZKEM1.Disconnect
    cmdClose.Enabled = False
    cmdOpen.Enabled = True
End Sub

Private Sub CZKEM1_OnAttTransaction(ByVal EnrollNumber As Long, ByVal IsInValid As Long, ByVal AttState As Long, ByVal VerifyMethod As Long, ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long)
    Text1.Text = Text1.Text & "Att Transaction: " & EnrollNumber & " " & AttState & " " & Year & "-" & _
        Month & "-" & Day & " " & Hour & ":" & Minute & ":" & Second & _
        Chr(13) & Chr(10)
End Sub

Private Sub CZKEM1_OnConnected()
    Text1.Text = Text1.Text & "Connected" & Chr(13) & Chr(10)
End Sub

Private Sub CZKEM1_OnDisConnected()
    Text1.Text = Text1.Text & "DisConnected" & Chr(13) & Chr(10)

End Sub

Private Sub CZKEM1_OnKeyPress(ByVal Key As Long)
    Text1.Text = Text1.Text & "Key: " & Key & Chr(13) & Chr(10)
End Sub

Private Sub CZKEM1_OnVerify(ByVal UserID As Long)
    Dim s As String
    If UserID > 0 Then
        MsgBox "Verified UserID=" + CStr(UserID)
    Else
        MsgBox "Verify Fail"
    
End Sub

Private Sub Form_Load()
    cmdClose.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CZKEM1.Disconnect
End Sub
