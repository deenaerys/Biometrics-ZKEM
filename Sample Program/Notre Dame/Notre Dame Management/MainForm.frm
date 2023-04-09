VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Begin VB.Form MainForm 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Notre Dame SchoolID Management Software"
   ClientHeight    =   3330
   ClientLeft      =   600
   ClientTop       =   4500
   ClientWidth     =   9120
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3210
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9015
      _Version        =   851968
      _ExtentX        =   15901
      _ExtentY        =   5662
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   1050
         Left            =   360
         TabIndex        =   1
         Top             =   675
         Width           =   2040
         _Version        =   851968
         _ExtentX        =   3598
         _ExtentY        =   1852
         _StockProps     =   79
         Caption         =   "Student's Remarks"
         TextAlignment   =   10
         Appearance      =   6
         Picture         =   "MainForm.frx":15162
         ImageAlignment  =   6
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   330
         Left            =   7245
         TabIndex        =   2
         Top             =   2295
         Width           =   1320
         _Version        =   851968
         _ExtentX        =   2328
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "E&xit"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   1050
         Left            =   2430
         TabIndex        =   3
         Top             =   675
         Width           =   2040
         _Version        =   851968
         _ExtentX        =   3598
         _ExtentY        =   1852
         _StockProps     =   79
         Caption         =   "General Announcement"
         TextAlignment   =   10
         Appearance      =   6
         Picture         =   "MainForm.frx":2A2D4
         ImageAlignment  =   6
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   1050
         Left            =   4500
         TabIndex        =   4
         Top             =   675
         Width           =   2040
         _Version        =   851968
         _ExtentX        =   3598
         _ExtentY        =   1852
         _StockProps     =   79
         Caption         =   "Server Settings"
         TextAlignment   =   10
         Appearance      =   6
         Picture         =   "MainForm.frx":3F446
         ImageAlignment  =   6
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   1050
         Left            =   6570
         TabIndex        =   5
         Top             =   675
         Width           =   2040
         _Version        =   851968
         _ExtentX        =   3598
         _ExtentY        =   1852
         _StockProps     =   79
         Caption         =   "Attendance Logs"
         TextAlignment   =   10
         Appearance      =   6
         Picture         =   "MainForm.frx":404D8
         ImageAlignment  =   6
         TextImageRelation=   1
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub PushButton1_Click()
    frmRemarks.Show vbModal
End Sub

Private Sub PushButton2_Click()
    frmReader.Show vbModal
End Sub

Private Sub PushButton3_Click()
    frmCamera.Show vbModal
End Sub

Private Sub PushButton4_Click()
    frmLogs.Show vbModal
End Sub

Private Sub PushButton5_Click()
    End
End Sub
