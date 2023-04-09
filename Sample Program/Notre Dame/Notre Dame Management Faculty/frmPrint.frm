VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print Report"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   495
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.RadioButton rdDate 
      Height          =   330
      Left            =   630
      TabIndex        =   2
      Top             =   360
      Width           =   3705
      _Version        =   851968
      _ExtentX        =   6535
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Print Report By Date"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      Appearance      =   6
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   330
      Left            =   3060
      TabIndex        =   0
      Top             =   1440
      Width           =   1320
      _Version        =   851968
      _ExtentX        =   2328
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Print"
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
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   330
      Left            =   1710
      TabIndex        =   1
      Top             =   1440
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
   Begin XtremeSuiteControls.RadioButton rdStudent 
      Height          =   330
      Left            =   630
      TabIndex        =   3
      Top             =   675
      Width           =   3705
      _Version        =   851968
      _ExtentX        =   6535
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Print Report By Student"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      Appearance      =   6
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PushButton1_Click()
    If rdDate.Value = True Then
         With frmLogs.CrystalReport1
                 .DiscardSavedData = True
                 .ReportFileName = GetAppPath & "rptLogsByDate.rpt"
                 .DataFiles(0) = GetAppPath & "school.mdb"
                 .Destination = crptToWindow
                 .Action = 1
        End With
        Unload Me
    Else
         With frmLogs.CrystalReport1
                 .DiscardSavedData = True
                 .ReportFileName = GetAppPath & "rptLogsBystudent.rpt"
                 .DataFiles(0) = GetAppPath & "school.mdb"
                 .Destination = crptToWindow
                 .Action = 1
        End With
        Unload Me
    End If
    
    
End Sub

Private Sub PushButton3_Click()
    Unload Me
End Sub
