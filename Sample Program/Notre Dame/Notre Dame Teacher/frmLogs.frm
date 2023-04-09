VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmLogs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Logs"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   375
      Left            =   6840
      TabIndex        =   23
      Top             =   450
      Width           =   2175
      _Version        =   851968
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Show Homeowner's Detail"
      Appearance      =   6
   End
   Begin VSFlex7Ctl.VSFlexGrid VSF 
      Height          =   3975
      Left            =   45
      TabIndex        =   0
      Top             =   1305
      Width           =   10320
      _cx             =   18203
      _cy             =   7011
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14737632
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin XtremeSuiteControls.FlatEdit txtAC 
      Height          =   285
      Left            =   1845
      TabIndex        =   1
      Top             =   5850
      Width           =   3525
      _Version        =   851968
      _ExtentX        =   6218
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLName 
      Height          =   285
      Left            =   1845
      TabIndex        =   2
      Top             =   6480
      Width           =   3525
      _Version        =   851968
      _ExtentX        =   6218
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFName 
      Height          =   285
      Left            =   1845
      TabIndex        =   3
      Top             =   6795
      Width           =   3525
      _Version        =   851968
      _ExtentX        =   6218
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMName 
      Height          =   285
      Left            =   1845
      TabIndex        =   4
      Top             =   7110
      Width           =   3525
      _Version        =   851968
      _ExtentX        =   6218
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.TabControl TabControl5 
      Height          =   2580
      Left            =   7740
      TabIndex        =   5
      Top             =   5805
      Width           =   2580
      _Version        =   851968
      _ExtentX        =   4551
      _ExtentY        =   4551
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      PaintManager.ShowTabs=   0   'False
      Begin VB.Image ImageH 
         Height          =   2475
         Left            =   45
         Stretch         =   -1  'True
         Top             =   45
         Width           =   2475
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtNotes 
      Height          =   555
      Left            =   1845
      TabIndex        =   6
      Top             =   7740
      Width           =   5685
      _Version        =   851968
      _ExtentX        =   10028
      _ExtentY        =   979
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtGender 
      Height          =   285
      Left            =   1845
      TabIndex        =   14
      Top             =   7425
      Width           =   3525
      _Version        =   851968
      _ExtentX        =   6218
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtSearch 
      Height          =   330
      Left            =   1305
      TabIndex        =   16
      Top             =   135
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton PushButton5 
      Height          =   330
      Left            =   6885
      TabIndex        =   17
      Top             =   135
      Width           =   1365
      _Version        =   851968
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Find"
      Appearance      =   1
      Picture         =   "frmLogs.frx":0000
   End
   Begin XtremeSuiteControls.DateTimePicker dtFrom 
      Height          =   330
      Left            =   4500
      TabIndex        =   19
      Top             =   135
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   582
      _StockProps     =   68
      CustomFormat    =   "mm/dd/yyyy HH:MM"
      Format          =   3
      CurrentDate     =   40675.4609259259
   End
   Begin XtremeSuiteControls.DateTimePicker dtTo 
      Height          =   330
      Left            =   4500
      TabIndex        =   21
      Top             =   495
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   582
      _StockProps     =   68
      CustomFormat    =   "mm/dd/yyyy HH:MM"
      Format          =   3
      CurrentDate     =   40675.4609259259
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   285
      Left            =   1845
      TabIndex        =   24
      Top             =   6165
      Width           =   3525
      _Version        =   851968
      _ExtentX        =   6218
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   195
      Left            =   360
      TabIndex        =   25
      Top             =   6210
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Student No.:"
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
   End
   Begin XtremeSuiteControls.Label Label15 
      Height          =   195
      Left            =   3195
      TabIndex        =   22
      Top             =   540
      Width           =   1275
      _Version        =   851968
      _ExtentX        =   2249
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "To:"
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
   End
   Begin XtremeSuiteControls.Label Label14 
      Height          =   195
      Left            =   3195
      TabIndex        =   20
      Top             =   180
      Width           =   1275
      _Version        =   851968
      _ExtentX        =   2249
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "From:"
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
   End
   Begin XtremeSuiteControls.Label Label13 
      Height          =   195
      Left            =   45
      TabIndex        =   18
      Top             =   180
      Width           =   1275
      _Version        =   851968
      _ExtentX        =   2249
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "AC Number:"
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
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Left            =   45
      TabIndex        =   15
      Top             =   945
      Width           =   2715
      _Version        =   851968
      _ExtentX        =   4789
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Logs"
      ForeColor       =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   2430
      X2              =   10350
      Y1              =   5535
      Y2              =   5535
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   375
      Left            =   45
      TabIndex        =   13
      Top             =   5400
      Width           =   2310
      _Version        =   851968
      _ExtentX        =   4075
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Student's Profile"
      ForeColor       =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label9 
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   5895
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "AC Number:"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   6525
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Last Name:"
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
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   6840
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "First Name:"
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
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   7155
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Middle Name:"
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
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   7470
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Grade/Year:"
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
   End
   Begin XtremeSuiteControls.Label Label8 
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   7785
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Remarks:"
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
   End
End
Attribute VB_Name = "frmLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub SearchHomeOwners(strCriteria As String, strDatefrom As String, strDateto As String)
    Dim RSB As New ADODB.Recordset
    Set RSB = cnMSSQL.Execute("select acnumber, logdt, readerip from attlog where acnumber like '" & strSearch & "%" & "' and logdt >= '" & strDatefrom & "' and logdt <= '" & strDateto & "'")
    
    
    
    
    Set VSF.DataSource = RSB
    VSF.ColWidth(0) = VSF.Width / 4
    VSF.ColWidth(1) = VSF.Width / 4
    VSF.ColWidth(2) = VSF.Width / 4
    
    
    VSF.ColAlignment(0) = flexAlignLeftCenter
    VSF.ColAlignment(1) = flexAlignLeftCenter
    VSF.ColAlignment(2) = flexAlignLeftCenter
    
    
    
    VSF.TextMatrix(0, 0) = "ACNumber"
    VSF.TextMatrix(0, 1) = "Log Time"
    VSF.TextMatrix(0, 2) = "Reader IP"
        
End Sub

Private Sub Form_Load()
clearTxt
End Sub
Sub clearTxt()

    txtAC.Text = vbNullString
    txtLName.Text = vbNullString
    txtFName.Text = vbNullString
    txtMName.Text = vbNullString
    txtAge.Text = vbNullString
    txtGender.Text = vbNullString
    txtAddress.Text = vbNullString
    txtGender.Text = vbNullString
    txtNotes.Text = vbNullString
    txtTelephone.Text = vbNullString
    ImageH.Picture = LoadPicture()
    
End Sub


Private Sub PushButton1_Click()
    Dim FSO As New FileSystemObject
    Dim Rsj As New ADODB.Recordset
    On Error GoTo err_exit
    Set Rsj = cnMSSQL.Execute("select * from homeowners where acnumber = '" & VSF.TextMatrix(VSF.Row, 0) & "'")
    If Not Rsj.EOF Then
        txtAC.Text = Rsj!acnumber
        txtLName.Text = Rsj!lname
        txtFName.Text = Rsj!fname
        txtMName.Text = Rsj!mname
        txtAge.Text = Rsj!age
        txtAddress.Text = Rsj!Address
        txtGender.Text = Rsj!gender
        txtNotes.Text = Rsj!Notes
        txtTelephone.Text = Rsj!telephone
        strAC = Rsj!acnumber
        Set FSO = New FileSystemObject
        If FSO.FileExists(App.Path & "\Homeowners\" & Rsj!acnumber & ".jpg") Then
            ImageH.Picture = LoadPicture(App.Path & "\Homeowners\" & Rsj!acnumber & ".jpg")
            DD = App.Path & "\Homeowners\" & Rsj!acnumber & ".jpg"
        
        Else
            ImageH.Picture = LoadPicture(App.Path & "\Homeowners\nophoto.jpg")
            DD = ""
        End If
        AppendState = "Edit"
        
    
    Else
        AppendState = "None"
        strAC = ""
        clearTxt
    End If
    
err_exit:
    
End Sub



Private Sub PushButton5_Click()
    Dim XTo As String
    Dim XFrom As String
    If dtFrom.Value > dtTo.Value Then
        MsgBox "Starting date cannot be larger than ending date"
        Exit Sub
    End If
    XFrom = Format(dtFrom.Year & "-" & dtFrom.Month & "-" & dtFrom.Day & "  " & dtFrom.Hour & ":" & dtFrom.Minute & ":00", "yyyy-mm-dd HH:MM:SS")
    XTo = Format(dtTo.Year & "-" & dtTo.Month & "-" & dtTo.Day & "  " & dtTo.Hour & ":" & dtTo.Minute & ":00", "yyyy-mm-dd HH:MM:SS")
    SearchHomeOwners txtSearch.Text, XFrom, XTo
End Sub

Private Sub VSF_Click()
    Dim FSO As New FileSystemObject
    If VSF.TextMatrix(VSF.Row, 3) = "Exit" Then
        If FSO.FileExists(App.Path & "/Exit/" & VSF.TextMatrix(VSF.Row, 2)) Then
            Image1.Picture = LoadPicture(App.Path & "/Exit/" & VSF.TextMatrix(VSF.Row, 2))
        Else
            Image1.Picture = LoadPicture()
        End If
    ElseIf VSF.TextMatrix(VSF.Row, 3) = "Entrance" Then
        If FSO.FileExists(App.Path & "/Entrance/" & VSF.TextMatrix(VSF.Row, 2)) Then
            Image1.Picture = LoadPicture(App.Path & "/Entrance/" & VSF.TextMatrix(VSF.Row, 2))
        Else
            Image1.Picture = LoadPicture()
        End If
    End If
End Sub
