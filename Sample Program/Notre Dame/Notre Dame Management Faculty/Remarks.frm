VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmRemarks 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Students"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid VSF 
      Height          =   5820
      Left            =   45
      TabIndex        =   0
      Top             =   630
      Width           =   10455
      _cx             =   18441
      _cy             =   10266
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
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   330
      Left            =   9135
      TabIndex        =   1
      Top             =   270
      Width           =   1320
      _Version        =   851968
      _ExtentX        =   2328
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Remove"
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
      Left            =   7785
      TabIndex        =   2
      Top             =   270
      Width           =   1320
      _Version        =   851968
      _ExtentX        =   2328
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Add"
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
   Begin XtremeSuiteControls.FlatEdit txtSearch 
      Height          =   330
      Left            =   4230
      TabIndex        =   3
      Top             =   270
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
      TabIndex        =   4
      Top             =   270
      Width           =   870
      _Version        =   851968
      _ExtentX        =   1535
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Find"
      Appearance      =   1
      Picture         =   "Remarks.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cboSearch 
      Height          =   315
      Left            =   1260
      TabIndex        =   5
      Top             =   270
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
   Begin XtremeSuiteControls.Label Label9 
      Height          =   195
      Left            =   45
      TabIndex        =   7
      Top             =   315
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
   Begin XtremeSuiteControls.Label Label10 
      Height          =   195
      Left            =   2925
      TabIndex        =   6
      Top             =   315
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
End
Attribute VB_Name = "frmRemarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strAC As String
Public strSN As String
Public SS As String
Public DD As String

Private Sub Form_Load()
    PostHomeOwners
        cboSearch.AddItem "Student Number"
    cboSearch.AddItem "Full Name"

End Sub
Sub SearchHomeOwners(strSearch As String, strCriteria As String)
    Dim RSB As New ADODB.Recordset
    If strCriteria = "Student Number" Then
        '"SELECT remarks.remarksid, userinfo.SSN, userinfo.name, userinfo.Badgenumber, userinfo.title, remarks.dtfrom, remarks.dtto, remarks.remarks FROM Remarks INNER JOIN USERINFO ON Remarks.SSN = USERINFO.SSN "
        'Set RSB = cnMSSQL.Execute("select SSN, badgenumber, name, title from userinfo where ssn like '" & "%" & strSearch & "%" & "'")
        Set RSB = cnMSSQL.Execute("SELECT remarks.remarksid, userinfo.SSN, userinfo.name, userinfo.Badgenumber, userinfo.title, remarks.dtfrom, remarks.dtto, remarks.remarks FROM Remarks INNER JOIN USERINFO ON Remarks.SSN = USERINFO.SSN where userinfo.ssn like '" & "%" & strSearch & "%" & "'")
    ElseIf strCriteria = "Full Name" Then
        'Set RSB = cnMSSQL.Execute("select SSN, badgenumber, name, title from userinfo where name like '" & "%" & strSearch & "%" & "'")
        Set RSB = cnMSSQL.Execute("SELECT remarks.remarksid, userinfo.SSN, userinfo.name, userinfo.Badgenumber, userinfo.title, remarks.dtfrom, remarks.dtto, remarks.remarks FROM Remarks INNER JOIN USERINFO ON Remarks.SSN = USERINFO.SSN where userinfo.name like '" & "%" & strSearch & "%" & "'")
    End If
    Set VSF.DataSource = RSB
    
    VSF.ColAlignment(0) = flexAlignLeftCenter
    VSF.ColAlignment(1) = flexAlignLeftCenter
    VSF.ColAlignment(2) = flexAlignLeftCenter
    VSF.ColAlignment(3) = flexAlignLeftCenter
    VSF.ColAlignment(4) = flexAlignLeftCenter
    VSF.ColAlignment(5) = flexAlignLeftCenter
    VSF.ColAlignment(6) = flexAlignLeftCenter
    VSF.ColAlignment(7) = flexAlignLeftCenter
    
    VSF.TextMatrix(0, 0) = "Remarks ID"
    VSF.TextMatrix(0, 1) = "Student No."
    VSF.TextMatrix(0, 2) = "Full Name"
    VSF.TextMatrix(0, 3) = "AC Number"
    VSF.TextMatrix(0, 4) = "Grade/Year"
    VSF.TextMatrix(0, 5) = "From"
    VSF.TextMatrix(0, 6) = "To"
    VSF.TextMatrix(0, 7) = "Remarks"
        
End Sub

Sub PostHomeOwners()
    Dim RSB As New ADODB.Recordset
    'RsX.Open "SELECT remarks.remarksid, userinfo.SSN, userinfo.Badgenumber, userinfo.title, remarks.dtfrom, remarks.dtto, remarks.remarks FROM Remarks INNER JOIN USERINFO ON Remarks.SSN = USERINFO.SSN "
    Set RSB = cnMSSQL.Execute("SELECT remarks.remarksid, userinfo.SSN, userinfo.name, userinfo.Badgenumber, userinfo.title, remarks.dtfrom, remarks.dtto, remarks.remarks FROM Remarks INNER JOIN USERINFO ON Remarks.SSN = USERINFO.SSN order by remarksid desc")
    Set VSF.DataSource = RSB
    'VSF.ColWidth(0) = VSF.Width / 5
    'VSF.ColWidth(1) = VSF.Width / 5
    'VSF.ColWidth(2) = VSF.Width / 5
    'VSF.ColWidth(3) = VSF.Width / 5
    'VSF.ColWidth(4) = VSF.Width / 5
    'VSF.ColWidth(5) = VSF.Width / 5
    'VSF.ColWidth(6) = VSF.Width / 5
    'VSF.ColWidth(7) = VSF.Width / 5
    
    VSF.ColAlignment(0) = flexAlignLeftCenter
    VSF.ColAlignment(1) = flexAlignLeftCenter
    VSF.ColAlignment(2) = flexAlignLeftCenter
    VSF.ColAlignment(3) = flexAlignLeftCenter
    VSF.ColAlignment(4) = flexAlignLeftCenter
    VSF.ColAlignment(5) = flexAlignLeftCenter
    VSF.ColAlignment(6) = flexAlignLeftCenter
    VSF.ColAlignment(7) = flexAlignLeftCenter
    
    VSF.TextMatrix(0, 0) = "Remarks ID"
    VSF.TextMatrix(0, 1) = "Student No."
    VSF.TextMatrix(0, 2) = "Full Name"
    VSF.TextMatrix(0, 3) = "AC Number"
    VSF.TextMatrix(0, 4) = "Grade/Year"
    VSF.TextMatrix(0, 5) = "From"
    VSF.TextMatrix(0, 6) = "To"
    VSF.TextMatrix(0, 7) = "Remarks"
        
End Sub

Private Sub PushButton1_Click()
    Dim rsp As New ADODB.Recordset
    If strSN = "" Then
        MsgBox "Please select remarks to delete.", vbExclamation, Me.Caption
        Exit Sub
    Else
        If MsgBox("Are you sure you want to delete the selected remarks?", vbYesNo + vbQuestion, Me.Caption) = vbNo Then Exit Sub
        Set rsp = cnMSSQL.Execute("delete from remarks where remarksid = '" & strSN & "'")
        PostHomeOwners
        strSN = ""
    End If
End Sub

Private Sub PushButton2_Click()
    'If strSN = "" Then
    '    MsgBox
    frmSelectStudent.Show vbModal
End Sub

Private Sub PushButton5_Click()
    SearchHomeOwners txtSearch.Text, cboSearch.Text
End Sub

Private Sub VSF_Click()
    On Error GoTo err1
    
    strSN = VSF.TextMatrix(VSF.Row, 0)
    Exit Sub
err1:
    strSN = ""
End Sub

Private Sub VSF_SelChange()
VSF_Click
End Sub

