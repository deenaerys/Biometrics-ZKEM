VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Begin VB.Form frmReader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scrolling Text"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6765
   Icon            =   "frmReader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   330
      Left            =   5175
      TabIndex        =   2
      Top             =   1800
      Width           =   1320
      _Version        =   851968
      _ExtentX        =   2328
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Update"
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
   Begin XtremeSuiteControls.FlatEdit txtIPEntrance 
      Height          =   1095
      Left            =   1620
      TabIndex        =   0
      Top             =   315
      Width           =   4875
      _Version        =   851968
      _ExtentX        =   8599
      _ExtentY        =   1931
      _StockProps     =   77
      BackColor       =   -2147483643
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   330
      Left            =   3825
      TabIndex        =   3
      Top             =   1800
      Width           =   1320
      _Version        =   851968
      _ExtentX        =   2328
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Close"
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
   Begin VB.Line Line1 
      X1              =   225
      X2              =   6480
      Y1              =   1575
      Y2              =   1575
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1230
      _Version        =   851968
      _ExtentX        =   2170
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "ScrollingText:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim RSC As New ADODB.Recordset
    
    Set RSC = cnMSSQL.Execute("select * from bottomscroll")
    If Not RSC.EOF Then
        txtIPEntrance.Text = RSC!scrolltext
    End If
End Sub

Private Sub PushButton1_Click()
    If Trim$(txtIPEntrance.Text) = "" Then
        MsgBox "Please do not leave blank fields.", vbExclamation, "Mandatory Entry"
        Exit Sub
    End If
    
    cnMSSQL.Execute "update bottomscroll set scrolltext = '" & txtIPEntrance.Text & "'"

    MsgBox "Scrolling text changed.", vbInformation, Me.Caption
    Unload Me
End Sub

Private Sub PushButton2_Click()
    Unload Me
End Sub
