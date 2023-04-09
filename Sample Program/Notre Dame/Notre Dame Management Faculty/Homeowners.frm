VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmSelectStudent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Student"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.FlatEdit txtAC 
      Height          =   285
      Left            =   1350
      TabIndex        =   1
      Top             =   3465
      Width           =   4065
      _Version        =   851968
      _ExtentX        =   7170
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VSFlex7Ctl.VSFlexGrid VSF 
      Height          =   2805
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   7170
      _cx             =   12647
      _cy             =   4948
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
   Begin XtremeSuiteControls.FlatEdit txtLName 
      Height          =   285
      Left            =   1350
      TabIndex        =   2
      Top             =   4095
      Width           =   4065
      _Version        =   851968
      _ExtentX        =   7170
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
      Left            =   4230
      TabIndex        =   6
      Top             =   90
      Width           =   2130
      _Version        =   851968
      _ExtentX        =   3757
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton PushButton5 
      Height          =   330
      Left            =   6390
      TabIndex        =   9
      Top             =   90
      Width           =   870
      _Version        =   851968
      _ExtentX        =   1535
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Find"
      Appearance      =   1
      Picture         =   "Homeowners.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtStudNo 
      Height          =   285
      Left            =   1350
      TabIndex        =   10
      Top             =   3780
      Width           =   4065
      _Version        =   851968
      _ExtentX        =   7170
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtGY 
      Height          =   285
      Left            =   1350
      TabIndex        =   12
      Top             =   4410
      Width           =   4065
      _Version        =   851968
      _ExtentX        =   7170
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboSearch 
      Height          =   315
      Left            =   1260
      TabIndex        =   5
      Top             =   90
      Width           =   1545
      _Version        =   851968
      _ExtentX        =   2725
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.TabControl TabControl5 
      Height          =   1860
      Left            =   5490
      TabIndex        =   14
      Top             =   3420
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   3281
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      PaintManager.ShowTabs=   0   'False
      Begin VB.Image ImageH 
         Height          =   1755
         Left            =   45
         Stretch         =   -1  'True
         Top             =   45
         Width           =   1710
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtRemarks 
      Height          =   645
      Left            =   1350
      TabIndex        =   15
      Top             =   5490
      Width           =   4065
      _Version        =   851968
      _ExtentX        =   7170
      _ExtentY        =   1138
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   330
      Left            =   5985
      TabIndex        =   17
      Top             =   6435
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
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   330
      Left            =   4635
      TabIndex        =   18
      Top             =   6435
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
   Begin XtremeSuiteControls.DateTimePicker DtFrom 
      Height          =   285
      Left            =   1350
      TabIndex        =   19
      Top             =   4770
      Width           =   4020
      _Version        =   851968
      _ExtentX        =   7091
      _ExtentY        =   503
      _StockProps     =   68
      CustomFormat    =   "'Date:'ddd, dd MMM yyyy    'Time:' HH:mm"
      Format          =   3
      CurrentDate     =   40729.9911226852
   End
   Begin XtremeSuiteControls.DateTimePicker DtTo 
      Height          =   285
      Left            =   1350
      TabIndex        =   20
      Top             =   5130
      Width           =   4020
      _Version        =   851968
      _ExtentX        =   7091
      _ExtentY        =   503
      _StockProps     =   68
      CustomFormat    =   "'Date:'ddd, dd MMM yyyy    'Time:' HH:mm"
      Format          =   3
      CurrentDate     =   40729.9911226852
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   195
      Left            =   -135
      TabIndex        =   22
      Top             =   4815
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Date From:"
      ForeColor       =   192
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
      Left            =   -135
      TabIndex        =   21
      Top             =   5175
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Date To:"
      ForeColor       =   192
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
   Begin VB.Line Line2 
      X1              =   90
      X2              =   7245
      Y1              =   6390
      Y2              =   6390
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   195
      Left            =   -135
      TabIndex        =   16
      Top             =   5535
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Remarks:"
      ForeColor       =   192
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
   Begin XtremeSuiteControls.Label Label5 
      Height          =   195
      Left            =   -135
      TabIndex        =   13
      Top             =   4455
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
   Begin XtremeSuiteControls.Label Label11 
      Height          =   195
      Left            =   -135
      TabIndex        =   11
      Top             =   3825
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
   Begin XtremeSuiteControls.Label Label10 
      Height          =   195
      Left            =   2925
      TabIndex        =   8
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
   Begin XtremeSuiteControls.Label Label9 
      Height          =   195
      Left            =   45
      TabIndex        =   7
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   195
      Left            =   -135
      TabIndex        =   4
      Top             =   4140
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   195
      Left            =   -135
      TabIndex        =   3
      Top             =   3510
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
   Begin VB.Line Line1 
      X1              =   90
      X2              =   7245
      Y1              =   3330
      Y2              =   3330
   End
End
Attribute VB_Name = "frmSelectStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AppendState As String
Public strAC As String
Public strSN As String
Public SS As String
Public DD As String

Private Sub Form_Load()
    PostHomeOwners
    cboSearch.AddItem "Student Number"
    cboSearch.AddItem "Full Name"
    clearTxt
End Sub
Sub clearTxt()
    txtAC.Text = vbNullString
    txtLName.Text = vbNullString
    txtStudNo.Text = vbNullString
    txtGY.Text = vbNullString
    txtRemarks.Text = vbNullString
    dtFrom.Value = Now
    dtTo.Value = Now
    ImageH.Picture = LoadPicture()
End Sub
Sub SearchHomeOwners(strSearch As String, strCriteria As String)
    Dim RSB As New ADODB.Recordset
    If strCriteria = "Student Number" Then
        Set RSB = cnMSSQL.Execute("select SSN, badgenumber, name, title from userinfo where ssn like '" & "%" & strSearch & "%" & "'")
    ElseIf strCriteria = "Full Name" Then
        Set RSB = cnMSSQL.Execute("select SSN, badgenumber, name, title from userinfo where name like '" & "%" & strSearch & "%" & "'")
    End If
    Set VSF.DataSource = RSB
    VSF.ColWidth(0) = (VSF.Width / 2) / 3
    VSF.ColWidth(1) = (VSF.Width / 2) / 3
    VSF.ColWidth(2) = VSF.Width / 2
    VSF.ColWidth(3) = (VSF.Width / 2) / 3
    
    VSF.ColAlignment(0) = flexAlignLeftCenter
    VSF.ColAlignment(1) = flexAlignLeftCenter
    VSF.ColAlignment(2) = flexAlignLeftCenter
    VSF.ColAlignment(3) = flexAlignLeftCenter
    
    VSF.TextMatrix(0, 0) = "Student Number"
    VSF.TextMatrix(0, 1) = "AC Number"
    VSF.TextMatrix(0, 2) = "Full Name"
    VSF.TextMatrix(0, 3) = "Grade/Year"
        
End Sub

Sub PostHomeOwners()
    Dim RSB As New ADODB.Recordset
    Set RSB = cnMSSQL.Execute("select SSN, badgenumber, name, title from userinfo order by name")
    Set VSF.DataSource = RSB
    VSF.ColWidth(0) = (VSF.Width / 2) / 3
    VSF.ColWidth(1) = (VSF.Width / 2) / 3
    VSF.ColWidth(2) = VSF.Width / 2
    VSF.ColWidth(3) = (VSF.Width / 2) / 3
    
    VSF.ColAlignment(0) = flexAlignLeftCenter
    VSF.ColAlignment(1) = flexAlignLeftCenter
    VSF.ColAlignment(2) = flexAlignLeftCenter
    VSF.ColAlignment(3) = flexAlignLeftCenter
    
    VSF.TextMatrix(0, 0) = "Student Number"
    VSF.TextMatrix(0, 1) = "AC Number"
    VSF.TextMatrix(0, 2) = "Full Name"
    VSF.TextMatrix(0, 3) = "Grade/Year"
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmRemarks.strSN = ""
End Sub

Private Sub PushButton1_Click()
    If Trim$(txtStudNo.Text) = "" Then
        MsgBox "Please select student.", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    If dtTo.Value <= dtFrom.Value Then
        MsgBox "Date From should not be later than Date To.", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    If Trim$(txtRemarks.Text) = "" Then
        MsgBox "Please enter remarks field.", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    If MsgBox("Save remarks to selected student?", vbYesNo + vbQuestion, Me.Caption) = vbNo Then Exit Sub
    
    Dim rsy As New ADODB.Recordset
    Set rsy = cnMSSQL.Execute("insert into remarks (ssn, dtfrom, dtto, remarks) values ('" & txtStudNo.Text & "','" & dtFrom.Value & "', '" & dtTo.Value & "','" & txtRemarks.Text & "')")
    frmRemarks.PostHomeOwners
    If MsgBox("Remarks updates!" & vbCrLf & vbCrLf & "Would you like to add another?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
        clearTxt
    Else
        Unload Me
        frmRemarks.strSN = ""
    End If

End Sub

Private Sub PushButton2_Click()
    Unload Me
    frmRemarks.strSN = ""
End Sub

Private Sub PushButton5_Click()
    SearchHomeOwners txtSearch.Text, cboSearch.Text
End Sub
Private Sub VSF_Click()
    Dim fso As New FileSystemObject
    Dim Rsj As New ADODB.Recordset
    'On Error GoTo err_exit
    Set Rsj = cnMSSQL.Execute("select SSN, badgenumber, name, title from userinfo where ssn = '" & VSF.TextMatrix(VSF.Row, 0) & "'")
    If Not Rsj.EOF Then
        txtAC.Text = Rsj!badgenumber
        txtLName.Text = Rsj!Name
        txtStudNo.Text = Rsj!ssn
        txtGY.Text = Rsj!Title
        Set fso = New FileSystemObject
        If fso.FileExists(App.Path & "\Users\" & Rsj!ssn & ".jpg") Then
            ImageH.Picture = LoadPicture(App.Path & "\Users\" & Rsj!ssn & ".jpg")
        Else
            ImageH.Picture = LoadPicture(App.Path & "\Users\nophoto.jpg")
        End If
    Else
        clearTxt
    End If
End Sub

Private Sub VSF_SelChange()
VSF_Click
End Sub
