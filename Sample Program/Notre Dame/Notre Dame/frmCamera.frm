VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Begin VB.Form frmCamera 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Server"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   330
      Left            =   3825
      TabIndex        =   8
      Top             =   2115
      Width           =   1320
      _Version        =   851968
      _ExtentX        =   2328
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Update"
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
      Picture         =   "frmCamera.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtServer 
      Height          =   285
      Left            =   1710
      TabIndex        =   0
      Top             =   315
      Width           =   3390
      _Version        =   851968
      _ExtentX        =   5980
      _ExtentY        =   503
      _StockProps     =   77
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDatabase 
      Height          =   285
      Left            =   1710
      TabIndex        =   1
      Top             =   630
      Width           =   3390
      _Version        =   851968
      _ExtentX        =   5980
      _ExtentY        =   503
      _StockProps     =   77
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUID 
      Height          =   285
      Left            =   1710
      TabIndex        =   2
      Top             =   945
      Width           =   3390
      _Version        =   851968
      _ExtentX        =   5980
      _ExtentY        =   503
      _StockProps     =   77
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPWD 
      Height          =   285
      Left            =   1710
      TabIndex        =   3
      Top             =   1260
      Width           =   3390
      _Version        =   851968
      _ExtentX        =   5980
      _ExtentY        =   503
      _StockProps     =   77
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   330
      Left            =   1125
      TabIndex        =   9
      Top             =   2115
      Width           =   1320
      _Version        =   851968
      _ExtentX        =   2328
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Close"
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
      Picture         =   "frmCamera.frx":059A
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   330
      Left            =   2475
      TabIndex        =   10
      Top             =   2115
      Width           =   1320
      _Version        =   851968
      _ExtentX        =   2328
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Test"
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
      Picture         =   "frmCamera.frx":0B34
   End
   Begin VB.Line Line1 
      X1              =   540
      X2              =   5085
      Y1              =   1710
      Y2              =   1710
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   195
      Left            =   -315
      TabIndex        =   7
      Top             =   1305
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Password:"
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
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   195
      Left            =   -315
      TabIndex        =   6
      Top             =   990
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Username:"
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
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   195
      Left            =   -315
      TabIndex        =   5
      Top             =   675
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Database:"
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
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   195
      Left            =   -315
      TabIndex        =   4
      Top             =   360
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Server:"
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
      Alignment       =   1
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim RsC As New ADODB.Recordset
    
    Set RsC = cnMS.Execute("select * from msserver")
    If Not RsC.EOF Then
        txtServer.Text = RsC!servername
        txtDatabase.Text = RsC!serverdatabase
        txtUID.Text = RsC!serveruid
        txtPWD.Text = RsC!serverpwd
    End If
End Sub

Private Sub PushButton1_Click()
    If Trim$(txtServer.Text) = "" Or Trim$(txtDatabase.Text) = "" Or Trim$(txtUID.Text) = "" Or Trim$(txtPWD.Text) = "" Then
        MsgBox "Please do not leave blank fields.", vbExclamation, "Mandatory Entry"
        Exit Sub
    End If
    
    If testSQL(txtServer.Text, txtDatabase.Text, txtUID.Text, txtPWD.Text) Then
        'MsgBox "Connection Succeeded!", vbInformation, Me.Caption
        Set cnMSSQL = New ADODB.Connection
        cnMSSQL.CursorLocation = 1
        cnMSSQL.Open "Driver={SQL Server};" & _
               "Server=" & txtServer.Text & ";" & _
               "Database=" & txtDatabase.Text & ";" & _
               "Uid=" & txtUID.Text & ";" & _
               "Pwd=" & txtPWD.Text & ";"
        cnMS.Execute "update msserver set servername = '" & txtServer.Text & "', serverdatabase = '" & txtDatabase.Text & "', serveruid = '" & txtUID.Text & "', serverpwd = '" & txtPWD.Text & "'"

        
    Else
            If MsgBox("Connection Failed! Would you like to save it anyway?", vbCritical, Me.Caption) = vbYes Then
                cnMS.Execute "update msserver set servername = '" & txtServer.Text & "', serverdatabase = '" & txtDatabase.Text & "', serveruid = '" & txtUID.Text & "', serverpwd = '" & txtPWD.Text & "'"
            
            End If
    End If
    
    

    MsgBox "New server settings updated.", vbInformation, Me.Caption
    Unload Me
    
End Sub

Private Sub PushButton2_Click()
    Unload Me
End Sub

Private Sub PushButton3_Click()
    If Trim$(txtServer.Text) = "" Or Trim$(txtDatabase.Text) = "" Or Trim$(txtUID.Text) = "" Or Trim$(txtPWD.Text) = "" Then
        MsgBox "Please do not leave blank fields.", vbExclamation, "Mandatory Entry"
        Exit Sub
    End If
    
    If testSQL(txtServer.Text, txtDatabase.Text, txtUID.Text, txtPWD.Text) Then
        MsgBox "Connection Succeeded!", vbInformation, Me.Caption
    Else
            MsgBox "Connection Failed!", vbCritical, Me.Caption
    End If
End Sub
