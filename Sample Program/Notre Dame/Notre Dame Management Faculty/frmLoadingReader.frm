VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Begin VB.Form frmLoadingReader 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9660
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoadingReader.frx":0000
   ScaleHeight     =   1830
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   180
      Top             =   135
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   780
      Left            =   90
      TabIndex        =   0
      Top             =   495
      Width           =   9555
      _Version        =   851968
      _ExtentX        =   16854
      _ExtentY        =   1376
      _StockProps     =   79
      Caption         =   "Loading Readers..."
      ForeColor       =   65535
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmLoadingReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    MainForm.ZKEntrance.Disconnect
    MainForm.ZKExit.Disconnect

    Me.Refresh
    Dim RsC As New ADODB.Recordset
    Set RsC = cnMS.Execute("select * from readers")
    If Not RsC.EOF Then
        MainForm.StrReader1 = RsC!IPaddressentrance
        MainForm.StrReader2 = RsC!IPaddressexit
        MainForm.strCommkey1 = RsC!commkeyentrance
        MainForm.strCommkey2 = RsC!commkeyexit
        strIP1 = RsC!IPaddressentrance
        strIP2 = RsC!IPaddressexit
        
        Label1.Caption = "Connecting Reader 1..."
        If CheckConnection(CStr((RsC!IPaddressentrance))) Then
            MainForm.ZKEntrance.SetCommPassword CLng(RsC!commkeyentrance)
            If MainForm.ZKEntrance.Connect_Net(CStr(RsC!IPaddressentrance), 4370) Then
                MainForm.ZKEntrance.BASE64 = 0
                MainForm.ZKEntrance.RegEvent 1, 32767
                
                Me.Refresh
                Label1.Caption = "Reader 1 Connected..."
                MainForm.Reader1.Caption = "Reader 1: Connected..."
                MainForm.Reader1.ForeColor = vbGreen
                MainForm.Reader1.Visible = False
                Me.Refresh
            Else
                Me.Refresh
                MainForm.Reader1.Caption = "Reader 1: Not Connected..."
                MainForm.Reader1.ForeColor = vbRed
                MainForm.Reader1.Visible = True
                Label1.Caption = "Connecting Reader 2..."
                'MsgBox "Entrance Reader failed to connect.", vbExclamation, "LGVHA"
                Me.Refresh
            End If
        Else
            
        End If
        Me.Refresh
        DoEvents
        Label1.Caption = "Connecting Reader 2..."
        Me.Refresh
                
        If CheckConnection((CStr(RsC!IPaddressexit))) Then
            MainForm.ZKExit.SetCommPassword CLng(RsC!commkeyexit)
            If MainForm.ZKExit.Connect_Net(CStr(RsC!IPaddressexit), 4370) Then
                MainForm.ZKExit.BASE64 = 0
                MainForm.ZKExit.RegEvent 1, 32767
                Me.Refresh
                Label1.Caption = "Reader 2 Connected..."
                MainForm.Reader2.Caption = "Reader 2: Connected..."
                MainForm.Reader2.ForeColor = vbGreen
                MainForm.Reader2.Visible = False
                Me.Refresh
            Else
                Me.Refresh
                MainForm.Reader2.Caption = "Reader 2: Not Connected..."
                MainForm.Reader2.ForeColor = vbRed
                MainForm.Reader2.Visible = True
                'MsgBox "Exit Reader failed to connect.", vbExclamation,
                Me.Refresh
            End If
        Else
        
        End If
    End If
Unload Me
End Sub
