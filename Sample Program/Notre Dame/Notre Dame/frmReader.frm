VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Begin VB.Form frmReader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reader Settings"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   330
      Left            =   5175
      TabIndex        =   8
      Top             =   2610
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
      Picture         =   "frmReader.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtIPEntrance 
      Height          =   285
      Left            =   2115
      TabIndex        =   0
      Top             =   315
      Width           =   3390
      _Version        =   851968
      _ExtentX        =   5980
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCommKeyEntrance 
      Height          =   285
      Left            =   2115
      TabIndex        =   1
      Top             =   630
      Width           =   3390
      _Version        =   851968
      _ExtentX        =   5980
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtIPExit 
      Height          =   285
      Left            =   2115
      TabIndex        =   2
      Top             =   1215
      Width           =   3390
      _Version        =   851968
      _ExtentX        =   5980
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCommKeyExit 
      Height          =   285
      Left            =   2115
      TabIndex        =   3
      Top             =   1530
      Width           =   3390
      _Version        =   851968
      _ExtentX        =   5980
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   330
      Left            =   3825
      TabIndex        =   9
      Top             =   2610
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
      Picture         =   "frmReader.frx":059A
   End
   Begin VB.Line Line1 
      X1              =   225
      X2              =   6480
      Y1              =   2385
      Y2              =   2385
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   1575
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Comm Key (2):"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   1260
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "IP Address (2):"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   675
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Comm Key(1):"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   360
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "IP Address(1):"
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
    Dim RsC As New ADODB.Recordset
    
    Set RsC = cnMS.Execute("select * from readers")
    If Not RsC.EOF Then
        txtIPEntrance.Text = RsC!IPaddressentrance
        txtCommKeyEntrance.Text = RsC!commkeyentrance
        txtIPExit.Text = RsC!IPaddressexit
        txtCommKeyExit.Text = RsC!commkeyexit
    End If
End Sub

Private Sub PushButton1_Click()
    If Trim$(txtIPEntrance.Text) = "" Or Trim$(txtCommKeyEntrance.Text) = "" Or Trim$(txtIPExit.Text) = "" Or Trim$(txtCommKeyExit.Text) = "" Then
        MsgBox "Please do not leave blank fields.", vbExclamation, "Mandatory Entry"
        Exit Sub
    End If
    
    cnMS.Execute "update readers set IPaddressentrance = '" & txtIPEntrance.Text & "', commkeyentrance = '" & txtCommKeyEntrance.Text & "', IPaddressexit = '" & txtIPExit.Text & "', commkeyexit = '" & txtCommKeyExit.Text & "'"

    MsgBox "New reader settings updated.", vbInformation, "Reader Settings"
    Unload Me
    frmLoadingReader.Show vbModal
End Sub

Private Sub PushButton2_Click()
    Unload Me
End Sub
