VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLogs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Logs"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8775
      Top             =   5940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   285
      Left            =   1845
      TabIndex        =   20
      Top             =   495
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MM/dd/yyyy HH:mm"
      Format          =   16580611
      CurrentDate     =   40731
   End
   Begin VSFlex7Ctl.VSFlexGrid VSF 
      Height          =   3975
      Left            =   0
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
      Cols            =   6
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
      Left            =   2250
      TabIndex        =   1
      Top             =   6030
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
      Left            =   2250
      TabIndex        =   2
      Top             =   6660
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
      Height          =   1815
      Left            =   6345
      TabIndex        =   3
      Top             =   5760
      Width           =   1905
      _Version        =   851968
      _ExtentX        =   3360
      _ExtentY        =   3201
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      PaintManager.ShowTabs=   0   'False
      Begin VB.Image ImageH 
         Height          =   1710
         Left            =   45
         Stretch         =   -1  'True
         Top             =   45
         Width           =   1800
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtGY 
      Height          =   285
      Left            =   2250
      TabIndex        =   8
      Top             =   6975
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
   Begin XtremeSuiteControls.PushButton PushButton5 
      Height          =   690
      Left            =   7560
      TabIndex        =   10
      Top             =   90
      Width           =   1365
      _Version        =   851968
      _ExtentX        =   2408
      _ExtentY        =   1217
      _StockProps     =   79
      Caption         =   "Find"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      Picture         =   "frmLogs.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtSSN 
      Height          =   285
      Left            =   2250
      TabIndex        =   13
      Top             =   6345
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
      Height          =   285
      Left            =   5490
      TabIndex        =   15
      Top             =   90
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboSearch 
      Height          =   315
      Left            =   1845
      TabIndex        =   16
      Top             =   90
      Width           =   1950
      _Version        =   851968
      _ExtentX        =   3440
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton cmdPrint 
      Height          =   690
      Left            =   8955
      TabIndex        =   19
      Top             =   90
      Width           =   1365
      _Version        =   851968
      _ExtentX        =   2408
      _ExtentY        =   1217
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
      Appearance      =   1
      Picture         =   "frmLogs.frx":059A
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   285
      Left            =   5490
      TabIndex        =   21
      Top             =   495
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MM/dd/yyyy HH:mm"
      Format          =   16580611
      CurrentDate     =   40731
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   195
      Left            =   4185
      TabIndex        =   18
      Top             =   135
      Width           =   1275
      _Version        =   851968
      _ExtentX        =   2249
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Search String:"
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
      Left            =   630
      TabIndex        =   17
      Top             =   135
      Width           =   1185
      _Version        =   851968
      _ExtentX        =   2090
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Search By:"
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
   Begin XtremeSuiteControls.Label Label7 
      Height          =   195
      Left            =   765
      TabIndex        =   14
      Top             =   6390
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
      Left            =   4185
      TabIndex        =   12
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
      Left            =   630
      TabIndex        =   11
      Top             =   540
      Width           =   1185
      _Version        =   851968
      _ExtentX        =   2090
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Left            =   45
      TabIndex        =   9
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
      TabIndex        =   7
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
      Left            =   765
      TabIndex        =   6
      Top             =   6075
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
      Left            =   765
      TabIndex        =   5
      Top             =   6705
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Full Name:"
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
      Left            =   765
      TabIndex        =   4
      Top             =   7020
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
End
Attribute VB_Name = "frmLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strStudNo As String

Sub SearchHomeOwners(strCriteria As String, strDatefrom As String, strDateto As String, strString As String)
    strStudNo = ""
    Dim RSB As New ADODB.Recordset
    Dim RSC As New ADODB.Recordset
    
'    Set RSB = cnMSSQL.Execute("SELECT userinfo.SSN, userinfo.name, userinfo.title, userinfo.Badgenumber, CHECKINOUT.CHECKTIME,  CHECKINOUT.SENSORID, Machines.MachineAlias FROM CHECKINOUT INNER JOIN USERINFO ON CHECKINOUT.USERID = USERINFO.USERID inner join Machines on checkinout.sensorid = machines.machinenumber where checkinout.CHECKTIME >= '" & strDatefrom & "' and CHECKTIME <= '" & strDateto & "' order by CHECKTIME")
    
    If strCriteria = "All" Then
        'Set RSB = cnMSSQL.Execute("SELECT userinfo.SSN, userinfo.name, userinfo.title, userinfo.Badgenumber, CHECKINOUT.CHECKTIME,  CHECKINOUT.SENSORID FROM CHECKINOUT INNER JOIN USERINFO ON CHECKINOUT.USERID = USERINFO.USERID where checkinout.CHECKTIME >= '" & strDatefrom & "' and CHECKTIME <= '" & strDateto & "' order by CHECKTIME")
        Set RSB = cnMSSQL.Execute("SELECT userinfo.SSN, userinfo.name, userinfo.title, userinfo.Badgenumber, CHECKINOUT.CHECKTIME,  Machines.MachineAlias FROM CHECKINOUT INNER JOIN USERINFO ON CHECKINOUT.USERID = USERINFO.USERID inner join Machines on checkinout.sensorid = machines.machinenumber where checkinout.CHECKTIME >= '" & strDatefrom & "' and CHECKTIME <= '" & strDateto & "' order by CHECKTIME")
    ElseIf strCriteria = "Full Name" Then
        'Set RSB = cnMSSQL.Execute("SELECT userinfo.SSN, userinfo.name, userinfo.title, userinfo.Badgenumber, CHECKINOUT.CHECKTIME,  CHECKINOUT.SENSORID FROM CHECKINOUT INNER JOIN USERINFO ON CHECKINOUT.USERID = USERINFO.USERID where checkinout.CHECKTIME >= '" & strDatefrom & "' and CHECKTIME <= '" & strDateto & "' and userinfo.name like '" & "%" & strString & "%" & "'order by CHECKTIME")
        Set RSB = cnMSSQL.Execute("SELECT userinfo.SSN, userinfo.name, userinfo.title, userinfo.Badgenumber, CHECKINOUT.CHECKTIME, Machines.MachineAlias FROM CHECKINOUT INNER JOIN USERINFO ON CHECKINOUT.USERID = USERINFO.USERID inner join Machines on checkinout.sensorid = machines.machinenumber where checkinout.CHECKTIME >= '" & strDatefrom & "' and CHECKTIME <= '" & strDateto & "' and userinfo.name like '" & "%" & strString & "%" & "'order by CHECKTIME")
    ElseIf strCriteria = "Student Number" Then
        'Set RSB = cnMSSQL.Execute("SELECT userinfo.SSN, userinfo.name, userinfo.title, userinfo.Badgenumber, CHECKINOUT.CHECKTIME,  CHECKINOUT.SENSORID FROM CHECKINOUT INNER JOIN USERINFO ON CHECKINOUT.USERID = USERINFO.USERID where checkinout.CHECKTIME >= '" & strDatefrom & "' and CHECKTIME <= '" & strDateto & "' and userinfo.SSN like '" & "%" & strString & "%" & "'order by CHECKTIME")
        Set RSB = cnMSSQL.Execute("SELECT userinfo.SSN, userinfo.name, userinfo.title, userinfo.Badgenumber, CHECKINOUT.CHECKTIME, Machines.MachineAlias FROM CHECKINOUT INNER JOIN USERINFO ON CHECKINOUT.USERID = USERINFO.USERID inner join Machines on checkinout.sensorid = machines.machinenumber where checkinout.CHECKTIME >= '" & strDatefrom & "' and CHECKTIME <= '" & strDateto & "' and userinfo.SSN like '" & "%" & strString & "%" & "'order by CHECKTIME")
    End If
    
    Set RSC = cnMS.Execute("delete from logs")
    If Not RSB.EOF Then
        Do While Not RSB.EOF
            Set RSC = cnMS.Execute("INSERT INTO LOGS values ('" & RSB!ssn & "','" & RSB!Name & "','" & RSB!Title & "','" & RSB!badgenumber & "','" & RSB!checktime & "','" & RSB!machinealias & "')")
            RSB.MoveNext
        Loop
        cmdPrint.Enabled = True
    Else
        cmdPrint.Enabled = False
    End If
    
    
    
    Set VSF.DataSource = RSB
    
    VSF.ColWidth(0) = VSF.Width / 8
    VSF.ColWidth(1) = VSF.Width / 4
    VSF.ColWidth(2) = VSF.Width / 8
    VSF.ColWidth(3) = VSF.Width / 8
    VSF.ColWidth(4) = VSF.Width / 4
    VSF.ColWidth(5) = VSF.Width / 8
    
    VSF.ColAlignment(0) = flexAlignLeftCenter
    VSF.ColAlignment(1) = flexAlignLeftCenter
    VSF.ColAlignment(2) = flexAlignLeftCenter
    VSF.ColAlignment(3) = flexAlignLeftCenter
    VSF.ColAlignment(4) = flexAlignLeftCenter
    VSF.ColAlignment(5) = flexAlignLeftCenter
    
    
    VSF.TextMatrix(0, 0) = "Student No."
    VSF.TextMatrix(0, 1) = "Full Name"
    VSF.TextMatrix(0, 2) = "Grade/Year"
    VSF.TextMatrix(0, 3) = "AC Number"
    VSF.TextMatrix(0, 4) = "Date/Time"
    VSF.TextMatrix(0, 5) = "Reader"
    


End Sub

Private Sub cmdPrint_Click()
    frmPrint.Show vbModal
End Sub

Private Sub Form_Load()
Dim dtNow As String

    clearTxt
    cboSearch.AddItem "All"
    cboSearch.AddItem "Student Number"
    cboSearch.AddItem "Full Name"
    cboSearch.Text = "All"
    DtFrom.Value = Month(Now) & "/" & Day(Now) & "/" & Year(Now) & " 00:00"
    DtTo.Value = Now
End Sub
Sub clearTxt()
    VSF.ColWidth(0) = VSF.Width / 8
    VSF.ColWidth(1) = VSF.Width / 4
    VSF.ColWidth(2) = VSF.Width / 8
    VSF.ColWidth(3) = VSF.Width / 8
    VSF.ColWidth(4) = VSF.Width / 4
    VSF.ColWidth(5) = VSF.Width / 8
    
    VSF.ColAlignment(0) = flexAlignLeftCenter
    VSF.ColAlignment(1) = flexAlignLeftCenter
    VSF.ColAlignment(2) = flexAlignLeftCenter
    VSF.ColAlignment(3) = flexAlignLeftCenter
    VSF.ColAlignment(4) = flexAlignLeftCenter
    VSF.ColAlignment(5) = flexAlignLeftCenter
    
    
    VSF.TextMatrix(0, 0) = "Student No."
    VSF.TextMatrix(0, 1) = "Full Name"
    VSF.TextMatrix(0, 2) = "Grade/Year"
    VSF.TextMatrix(0, 3) = "AC Number"
    VSF.TextMatrix(0, 4) = "Date/Time"
    VSF.TextMatrix(0, 5) = "Reader ID"
    
    txtAC.Text = vbNullString
    txtLName.Text = vbNullString
    txtSSN.Text = vbNullString
    txtGY.Text = vbNullString
    ImageH.Picture = LoadPicture()
    
    cmdPrint.Enabled = False
End Sub
Private Sub PushButton5_Click()
    Dim XTo As String
    Dim XFrom As String
    If DtFrom.Value > DtTo.Value Then
        MsgBox "Starting date cannot be larger than ending date"
        Exit Sub
    End If
    XFrom = Format(DtFrom.Value, "mm/dd/yyyy HH:MM:SS")
    XTo = Format(DtTo.Value, "mm/dd/yyyy HH:MM:SS")
    
    
    SearchHomeOwners cboSearch.Text, XFrom, XTo, txtSearch.Text
End Sub
Private Sub VSF_Click()
    Dim fso As New FileSystemObject
    
    
    On Error GoTo err1
    
    strStudNo = VSF.TextMatrix(VSF.Row, 0)
    
    txtAC.Text = VSF.TextMatrix(VSF.Row, 3)
    txtLName.Text = VSF.TextMatrix(VSF.Row, 1)
    txtSSN.Text = VSF.TextMatrix(VSF.Row, 0)
    txtGY.Text = VSF.TextMatrix(VSF.Row, 2)
    
    Set fso = New FileSystemObject
    If fso.FileExists(App.Path & "\Users\" & strStudNo & ".jpg") Then
        ImageH.Picture = LoadPicture(App.Path & "\Users\" & strStudNo & ".jpg")
    Else
        
        If Not strStudNo = "" Then ImageH.Picture = LoadPicture(App.Path & "\Users\nophoto.jpg")
    End If
    
    
    
    Exit Sub
err1:
    strSN = ""
End Sub

Private Sub VSF_SelChange()
VSF_Click
End Sub

