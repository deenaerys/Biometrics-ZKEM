VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Begin VB.Form frmLoading 
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
   Picture         =   "frmLoading.frx":0000
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
      Caption         =   "Loading Entrance Camera..."
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
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Me.Refresh
    Dim RsC As New ADODB.Recordset
    Set RsC = cnMSSQL.Execute("select * from camera")
    If Not RsC.EOF Then
        Label1.Caption = "Loading Entrance Camera..."
        MainForm.EntranceLOG = MainForm.Cam1.Login(CStr(RsC!IPaddress), CLng(RsC!camport), CStr(RsC!camuser), CStr(RsC!campwd))
        If MainForm.EntranceLOG <> -1 Then
            MainForm.Cam1.StartRealPlay CLng(RsC!camentrance), 0, 0
            MainForm.Cam1Channel = CLng(RsC!camentrance)
        Else
            MsgBox "Entrance Camera failed to connect.", vbExclamation, "LGVHA"
            Label1.Caption = "Loading Exit Camera..."
        End If
        Me.Refresh
        DoEvents
        Label1.Caption = "Loading Exit Camera..."
        MainForm.ExitLOG = MainForm.Cam2.Login(CStr(RsC!IPaddress), CLng(RsC!camport), CStr(RsC!camuser), CStr(RsC!campwd))
        If MainForm.ExitLOG <> -1 Then
            MainForm.Cam2.StartRealPlay CLng(RsC!camexit), 0, 0
            MainForm.Cam2Channel = CLng(RsC!camexit)
        Else
            MsgBox "Exit Camera failed to connect.", vbExclamation, "LGVHA"
        End If
    End If
Unload Me
End Sub
