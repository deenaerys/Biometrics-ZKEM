VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Olie.dll"
Object = "{15519D4E-5365-4718-B5E8-8F541C781688}#1.1#0"; "CoolXPPanel.ocx"
Begin VB.Form MainForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "NotreDame"
   ClientHeight    =   5955
   ClientLeft      =   840
   ClientTop       =   6300
   ClientWidth     =   8655
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   8625
      TabIndex        =   25
      Top             =   4860
      Width           =   8655
      Begin XtremeSuiteControls.Label lblTime 
         Height          =   555
         Left            =   7650
         TabIndex        =   27
         Top             =   0
         Width           =   6585
         _Version        =   851968
         _ExtentX        =   11615
         _ExtentY        =   979
         _StockProps     =   79
         Caption         =   "ddd"
         ForeColor       =   65535
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblDate 
         Height          =   555
         Left            =   135
         TabIndex        =   26
         Top             =   0
         Width           =   7080
         _Version        =   851968
         _ExtentX        =   12488
         _ExtentY        =   979
         _StockProps     =   79
         Caption         =   "ddd"
         ForeColor       =   65535
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"MainForm.frx":058A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   480
         Left            =   0
         TabIndex        =   24
         Top             =   585
         Width           =   23850
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   0
      Top             =   1935
   End
   Begin zkemkeeperCtl.CZKEM ZKExit 
      Height          =   195
      Left            =   0
      OleObjectBlob   =   "MainForm.frx":0614
      TabIndex        =   4
      Top             =   3420
      Visible         =   0   'False
      Width           =   375
   End
   Begin zkemkeeperCtl.CZKEM ZKEntrance 
      Height          =   195
      Left            =   0
      OleObjectBlob   =   "MainForm.frx":0638
      TabIndex        =   3
      Top             =   3195
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   0
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   2745
   End
   Begin XtremeSuiteControls.TabControl LTab 
      Height          =   4335
      Left            =   1350
      TabIndex        =   1
      Top             =   2025
      Width           =   4155
      _Version        =   851968
      _ExtentX        =   7329
      _ExtentY        =   7646
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      PaintManager.ShowTabs=   0   'False
      Begin VB.Image LImage 
         Height          =   4155
         Left            =   90
         Stretch         =   -1  'True
         Top             =   90
         Width           =   3975
      End
   End
   Begin XtremeSuiteControls.TabControl RTab 
      Height          =   4335
      Left            =   9135
      TabIndex        =   15
      Top             =   2025
      Width           =   4155
      _Version        =   851968
      _ExtentX        =   7329
      _ExtentY        =   7646
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      PaintManager.ShowTabs=   0   'False
      Begin VB.Image RImage 
         Height          =   4155
         Left            =   90
         Stretch         =   -1  'True
         Top             =   90
         Width           =   3975
      End
   End
   Begin CoolXPPanel.xpPanel xpPanel1 
      Align           =   1  'Align Top
      Height          =   1815
      Left            =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3201
      ColorStyle      =   99
      GradientDirection=   1
      GradientFrom    =   16384
      GradientTo      =   16761024
      Begin VB.Image Image1 
         Height          =   1965
         Left            =   0
         Picture         =   "MainForm.frx":065C
         Stretch         =   -1  'True
         Top             =   -90
         Width           =   29985
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4455
      Width           =   825
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Device"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4185
      Width           =   825
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Refresh"
      Height          =   195
      Left            =   225
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4725
      Width           =   825
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Server"
      Height          =   240
      Left            =   225
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3870
      Width           =   825
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   250
      Left            =   7875
      TabIndex        =   38
      Top             =   9045
      Width           =   1725
      _Version        =   851968
      _ExtentX        =   3043
      _ExtentY        =   441
      _StockProps     =   79
      Caption         =   "Last Login..."
      ForeColor       =   8421504
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   250
      Left            =   360
      TabIndex        =   37
      Top             =   8685
      Width           =   1725
      _Version        =   851968
      _ExtentX        =   3043
      _ExtentY        =   441
      _StockProps     =   79
      Caption         =   "Last Login..."
      ForeColor       =   8421504
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image RI2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   645
      Left            =   11880
      Stretch         =   -1  'True
      Top             =   9270
      Width           =   600
   End
   Begin VB.Image RI1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   645
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   9270
      Width           =   600
   End
   Begin XtremeSuiteControls.Label RP2 
      Height          =   645
      Left            =   12510
      TabIndex        =   36
      Top             =   9270
      Width           =   2175
      _Version        =   851968
      _ExtentX        =   3836
      _ExtentY        =   1138
      _StockProps     =   79
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label RP1 
      Height          =   645
      Left            =   8910
      TabIndex        =   35
      Top             =   9270
      Width           =   2175
      _Version        =   851968
      _ExtentX        =   3836
      _ExtentY        =   1138
      _StockProps     =   79
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label rbTime 
      Height          =   420
      Left            =   7650
      TabIndex        =   34
      Top             =   7515
      Visible         =   0   'False
      Width           =   1950
      _Version        =   851968
      _ExtentX        =   3440
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Time"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lbTime 
      Height          =   420
      Left            =   45
      TabIndex        =   33
      Top             =   7470
      Visible         =   0   'False
      Width           =   1950
      _Version        =   851968
      _ExtentX        =   3440
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Time"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LP2 
      Height          =   645
      Left            =   4590
      TabIndex        =   32
      Top             =   9270
      Width           =   2175
      _Version        =   851968
      _ExtentX        =   3836
      _ExtentY        =   1138
      _StockProps     =   79
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label LP1 
      Height          =   645
      Left            =   1575
      TabIndex        =   31
      Top             =   9270
      Width           =   2175
      _Version        =   851968
      _ExtentX        =   3836
      _ExtentY        =   1138
      _StockProps     =   79
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image LI2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   645
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   9270
      Width           =   600
   End
   Begin VB.Image LI1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   645
      Left            =   945
      Stretch         =   -1  'True
      Top             =   9270
      Width           =   600
   End
   Begin XtremeSuiteControls.Label lblSql 
      Height          =   195
      Left            =   6615
      TabIndex        =   30
      Top             =   1800
      Width           =   1680
      _Version        =   851968
      _ExtentX        =   2963
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Database Connection..."
      ForeColor       =   255
      BackColor       =   -2147483633
      Alignment       =   2
      Transparent     =   -1  'True
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblStud2 
      Height          =   420
      Left            =   7830
      TabIndex        =   23
      Top             =   6660
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Employee #:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label RStudNo 
      Height          =   420
      Left            =   9630
      TabIndex        =   22
      Top             =   6660
      Width           =   4605
      _Version        =   851968
      _ExtentX        =   8123
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "AC Number:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblName2 
      Height          =   420
      Left            =   7830
      TabIndex        =   21
      Top             =   7065
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Name:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label RLName 
      Height          =   735
      Left            =   9630
      TabIndex        =   20
      Top             =   7110
      Width           =   5550
      _Version        =   851968
      _ExtentX        =   9790
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "AC Number:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblGY2 
      Height          =   420
      Left            =   7830
      TabIndex        =   19
      Top             =   7830
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Designation:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label RYear 
      Height          =   420
      Left            =   9630
      TabIndex        =   18
      Top             =   7830
      Width           =   4605
      _Version        =   851968
      _ExtentX        =   8123
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "AC Number:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LBLRemarks2 
      Height          =   420
      Left            =   7830
      TabIndex        =   17
      Top             =   8190
      Visible         =   0   'False
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Remarks:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label RNote 
      Height          =   780
      Left            =   9630
      TabIndex        =   16
      Top             =   8235
      Visible         =   0   'False
      Width           =   5190
      _Version        =   851968
      _ExtentX        =   9155
      _ExtentY        =   1376
      _StockProps     =   79
      Caption         =   "AC Number:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label LNote 
      Height          =   780
      Left            =   2160
      TabIndex        =   14
      Top             =   8235
      Visible         =   0   'False
      Width           =   5190
      _Version        =   851968
      _ExtentX        =   9155
      _ExtentY        =   1376
      _StockProps     =   79
      Caption         =   "AC Number: "
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblRemarks1 
      Height          =   420
      Left            =   360
      TabIndex        =   13
      Top             =   8190
      Visible         =   0   'False
      Width           =   1725
      _Version        =   851968
      _ExtentX        =   3043
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Remarks:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LYear 
      Height          =   420
      Left            =   2160
      TabIndex        =   12
      Top             =   7830
      Width           =   4335
      _Version        =   851968
      _ExtentX        =   7646
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "AC Number:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblGY1 
      Height          =   420
      Left            =   360
      TabIndex        =   11
      Top             =   7830
      Width           =   1725
      _Version        =   851968
      _ExtentX        =   3043
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Designation:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LLName 
      Height          =   735
      Left            =   2115
      TabIndex        =   10
      Top             =   7110
      Width           =   5190
      _Version        =   851968
      _ExtentX        =   9155
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "AC Number:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblName1 
      Height          =   420
      Left            =   360
      TabIndex        =   9
      Top             =   7065
      Width           =   1770
      _Version        =   851968
      _ExtentX        =   3122
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Name:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LStudNo 
      Height          =   420
      Left            =   2115
      TabIndex        =   8
      Top             =   6660
      Width           =   4335
      _Version        =   851968
      _ExtentX        =   7646
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "AC Number:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblStud1 
      Height          =   420
      Left            =   360
      TabIndex        =   7
      Top             =   6660
      Width           =   1770
      _Version        =   851968
      _ExtentX        =   3122
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Employee #:"
      ForeColor       =   16576
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Reader2 
      Height          =   195
      Left            =   10260
      TabIndex        =   6
      Top             =   1800
      Width           =   1965
      _Version        =   851968
      _ExtentX        =   3466
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Reader 2: Not Connected..."
      ForeColor       =   255
      BackColor       =   -2147483633
      Alignment       =   2
      Transparent     =   -1  'True
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Reader1 
      Height          =   195
      Left            =   2340
      TabIndex        =   5
      Top             =   1800
      Width           =   1965
      _Version        =   851968
      _ExtentX        =   3466
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Reader 1: Not Connected..."
      ForeColor       =   255
      BackColor       =   -2147483633
      Alignment       =   2
      Transparent     =   -1  'True
      AutoSize        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      X1              =   7470
      X2              =   7515
      Y1              =   2070
      Y2              =   10305
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StrReader1 As String
Public StrReader2 As String
Public strCommkey1 As String
Public strCommkey2 As String
Public Reader1Stat As Boolean
Public Reader2Stat As Boolean
Public EntranceLOG As Integer
Public ExitLOG As Integer
Public Cam1Channel As Integer
Public Cam2Channel As Integer

Sub EntranceCapture(strFilename As String)
    
    Cam1.JPEGCapturePicture CLng(Cam1Channel + 1), 0, 2, App.Path & "\Entrance\"
    Dim SS As String
    Dim aa As String
    aa = App.Path & "\Entrance\JPEGCapture\*.jpeg"
    SS = Dir(aa)
    Do
    If Len(SS) = 0 Then Exit Do
    ''do something with ss
        'Debug.Print ss
        Image1.Picture = LoadPicture(App.Path & "\Entrance\JPEGCapture\" & SS)
        Name App.Path & "\Entrance\JPEGCapture\" & SS As App.Path & "\Entrance\JPEGCapture\" & strFilename
        FileCopy App.Path & "\Entrance\JPEGCapture\" & strFilename, App.Path & "\Entrance\" & strFilename
        Kill App.Path & "\Entrance\JPEGCapture\" & strFilename
    SS = Dir
    Loop
    
    
    
    
    'Debug.Print xx
End Sub
Sub ExitCapture(strFilename As String)
    
    Cam2.JPEGCapturePicture CLng(Cam2Channel + 1), 0, 2, App.Path & "\Exit\"
    Dim SS As String
    Dim aa As String
    aa = App.Path & "\Exit\JPEGCapture\*.jpeg"
    SS = Dir(aa)
    Do
    If Len(SS) = 0 Then Exit Do
    ''do something with ss
        'Debug.Print ss
        Image2.Picture = LoadPicture(App.Path & "\Exit\JPEGCapture\" & SS)
        Name App.Path & "\Exit\JPEGCapture\" & SS As App.Path & "\Exit\JPEGCapture\" & strFilename
        FileCopy App.Path & "\Exit\JPEGCapture\" & strFilename, App.Path & "\Exit\" & strFilename
        Kill App.Path & "\Exit\JPEGCapture\" & strFilename
    SS = Dir
    Loop
    
    
    
    
    'Debug.Print xx
End Sub


Private Sub Command1_Click()
    Homeowners.Show vbModal
    
End Sub

Private Sub Command2_Click()
    End
End Sub


Private Sub Command3_Click()
frmLoadingReader.Show vbModal
End Sub

Private Sub Command4_Click()
    frmReader.Show vbModal
    
End Sub

Private Sub Command5_Click()
    frmLogs.Show vbModal
End Sub

Private Sub Command6_Click()
    frmCamera.Show vbModal
End Sub

Private Sub Form_Load()
    'Form1.Show vbModal
    Dim rsk As New ADODB.Recordset
    Image1.Picture = LoadPicture(App.Path & "\header.jpg")
    ClearLeft
    ClearRight
    LblDate.Caption = Format(Now, "dddddd")
    lblTime.Caption = Format(Now, "H:MM:SS AM/PM")
    'If Not ConnectSQL Then
    '    lblSql.Caption = "No Database Connection."
    '    lblSql.ForeColor = vbRed
    '    lblSql.Visible = True
    'Else
    '    lblSql.Caption = "Connected to Database."
    '    lblSql.ForeColor = vbGreen
    '    lblSql.Visible = False
    '    Set rsk = cnMSSQL.Execute("select * from bottomscroll")
    '    If Not rsk.EOF Then
    '        Label1.Caption = rsk!scrolltext
    '    End If
    'End If
'    frmResize
End Sub
Sub ClearLeft()
    LStudNo.Caption = ""
    LLName.Caption = ""
    'LFname.Caption = ""
    'LMName.Caption = ""
    LYear.Caption = ""
    LNote.Caption = ""
    LImage.Picture = LoadPicture()
End Sub
Sub ClearRight()
    RStudNo.Caption = ""
    RLName.Caption = ""
    'RFName.Caption = ""
    'RMName.Caption = ""
    RYear.Caption = ""
    RNote.Caption = ""
    RImage.Picture = LoadPicture()
End Sub



Sub frmResize()
    Label2.Left = 360
    Label3.Left = (Me.Width / 2) + 360
    Label2.Top = Picture1.Top - 940
    Label3.Top = Picture1.Top - 940
        
    LI1.Top = Picture1.Top - 690
    LI2.Top = Picture1.Top - 690
    LP1.Top = Picture1.Top - 690
    LP2.Top = Picture1.Top - 690
    
    RI1.Top = Picture1.Top - 690
    RI2.Top = Picture1.Top - 690
    RP1.Top = Picture1.Top - 690
    RP2.Top = Picture1.Top - 690
    
    Reader2.Left = ((Me.Width / 4) * 3) - (Reader2.Width / 2)
    Reader1.Left = (Me.Width / 4) - (Reader1.Width / 2)
    lblSql.Left = (Me.Width / 2) - (lblSql.Width / 2)
    
    Line1.X1 = Me.Width / 2
    Line1.X2 = Me.Width / 2
    
    RTab.Left = ((Me.Width / 4) * 3) - (RTab.Width / 2)
    LTab.Left = (Me.Width / 4) - (LTab.Width / 2)
    
    lblTime.Left = ((Me.Width / 4) * 3) - (lblTime.Width / 2)
    LblDate.Left = (Me.Width / 4) - (LblDate.Width / 2)
    
    LblStud2.Left = (Me.Width / 2) + 360
    LblName2.Left = (Me.Width / 2) + 360
    LblGY2.Left = (Me.Width / 2) + 360
    LBLRemarks2.Left = (Me.Width / 2) + 360
    
    RStudNo.Left = (Me.Width / 2) + 360 + LblStud2.Width
    RLName.Left = (Me.Width / 2) + 360 + LblName2.Width
    RYear.Left = (Me.Width / 2) + 360 + LblGY2.Width
    RNote.Left = (Me.Width / 2) + 360 + LBLRemarks2.Width
    
    LI2.Left = (Me.Width / 4) + 100
    RI2.Left = ((Me.Width / 4) * 3) + 100
    
    LP2.Left = (Me.Width / 4) + 745
    RP2.Left = ((Me.Width / 4) * 3) + 745
    
    LI1.Left = (Me.Width / 4) - 2920
    RI1.Left = ((Me.Width / 4) * 3) - 2920
    
    LP1.Left = ((Me.Width / 4) - 2920) + 645
    RP1.Left = (((Me.Width / 4) * 3) - 2920) + 645
    
    Command2.Left = Me.Width / 4
    Command3.Left = Me.Width / 4
    Command4.Left = Me.Width / 4
    Command6.Left = Me.Width / 4
End Sub








Private Sub Image1_DblClick()
End
End Sub

Private Sub Timer1_Timer()
    Me.Refresh
    frmResize
    
    If Not ConnectSQL Then
        lblSql.Caption = "No Database Connection."
        lblSql.ForeColor = vbRed
        lblSql.Visible = True
    Else
        lblSql.Caption = "Connected to Database."
        lblSql.ForeColor = vbGreen
        lblSql.Visible = False
        Set rsk = cnMSSQL.Execute("select * from bottomscroll")
        If Not rsk.EOF Then
            Label1.Caption = rsk!scrolltext
        End If
    End If

    frmLoadingReader.Show vbModal
    Timer1.Enabled = False
    Me.WindowState = 2
    frmResize
End Sub

Private Sub Timer2_Timer()
    On Error GoTo err_timeR
    
    ZKEntrance.BASE64 = 0
    ZKEntrance.RegEvent 1, 32767
    
    ZKExit.BASE64 = 0
    ZKExit.RegEvent 1, 32767
    
    lngDiff = DateDiff("n", dtMRem, Now)
    lngDiffPIC1 = DateDiff("s", dtSRem1, Now)
    lngDiffPIC2 = DateDiff("s", dtSRem2, Now)
    
    If lngDiffPIC1 >= 10 Then
        If LStudNo.Caption = "" Then
        Else
            LMove
            ClearLeft
        End If
        LImage.Picture = LoadPicture()
        lngDiffPIC1 = 0
        dtSRem1 = Now
            
    End If
    If lngDiffPIC2 >= 10 Then
        If RStudNo.Caption = "" Then
        Else
            RMove
            ClearRight
        End If
        RImage.Picture = LoadPicture()
        lngDiffPIC2 = 0
        dtSRem2 = Now
            
    End If
    
    
    
    If lngDiff < 3 Then Exit Sub
'Check SQL
    Dim rsk As New ADODB.Recordset
    If Not ConnectSQL Then
        lblSql.Caption = "No Database Connection."
        lblSql.ForeColor = vbRed
        lblSql.Visible = True
    Else
        lblSql.Caption = "Connected to Database."
        lblSql.ForeColor = vbGreen
        lblSql.Visible = False
        Set rsk = cnMSSQL.Execute("select * from bottomscroll")
        If Not rsk.EOF Then
            Label1.Caption = rsk!scrolltext
        End If
        
        
        
    End If
'Check Reader
    If CheckConnection(CStr(StrReader1)) Then
        If Reader1Stat Then
            Reader1.Caption = "Reader 1: Connected..."
            Reader1.ForeColor = vbGreen
            Reader1.Visible = False
            ZKEntrance.RegEvent 1, 32767
        Else
            ZKEntrance.SetCommPassword CLng(strCommkey1)
            If ZKEntrance.Connect_Net(CStr(StrReader1), CLng(4370)) Then
                ZKEntrance.BASE64 = 0
                ZKEntrance.RegEvent 1, 32767
                Reader1.Caption = "Reader 1: Connected..."
                Reader1.ForeColor = vbGreen
                Reader1.Visible = False
                Reader1Stat = True
            Else
                Reader1.Caption = "Reader 1: Not Connected..."
                Reader1.ForeColor = vbRed
                Reader1.Visible = True
            End If
        End If
    Else
        Reader1.Caption = "Reader 1: Not Connected..."
        Reader1.ForeColor = vbRed
        Reader1.Visible = True
        Reader1Stat = False
        
    End If


    If CheckConnection(CStr(StrReader2)) Then
        If Reader2Stat Then
            Reader2.Caption = "Reader 2: Connected..."
            Reader2.ForeColor = vbGreen
            Reader2.Visible = False
            ZKExit.RegEvent 1, 32767
        Else
            ZKExit.SetCommPassword CLng(strCommkey2)
            If ZKExit.Connect_Net(CStr(StrReader2), CLng(4370)) Then
                ZKExit.BASE64 = 0
                ZKExit.RegEvent 1, 32767
                Reader2.Caption = "Reader 2: Connected..."
                Reader2.ForeColor = vbGreen
                Reader2.Visible = False
                Reader2Stat = True
            Else
                Reader2.Caption = "Reader 2: Not Connected..."
                Reader2.ForeColor = vbRed
                Reader2.Visible = True
            End If
        End If
    Else
        Reader2.Caption = "Reader 2: Not Connected..."
        Reader2.ForeColor = vbRed
        Reader2.Visible = True
        Reader2Stat = False
        
    End If
    lngDiff = 0
    dtMRem = Now
    Exit Sub
err_timeR:
    Call Main
End Sub

Private Sub Timer3_Timer()
    DoEvents
    LblDate.Caption = Format(Now, "dddddd")
    lblTime.Caption = Format(Now, "H:MM:SS AM/PM")
    Label1.Move Label1.Left - 25
    If Label1.Left < -Label1.Width Then Label1.Left = Picture1.Width
End Sub

Private Sub ZKEntrance_OnAttTransaction(ByVal EnrollNumber As Long, ByVal IsInValid As Long, ByVal AttState As Long, ByVal VerifyMethod As Long, ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long)
    Dim Xdt As String
    Dim strFname As String
    Dim rsL As New ADODB.Recordset
    
    Xdt = Format(Month & "/" & Day & "/" & Year & "  " & Hour & ":" & Minute & ":" & Second, "mm/dd/yyyy HH:MM:SS AM/PM")
'    Xdt = Format(Hour & ":" & Minute & ":" & Second, "HH:MM:SS AM/PM")
    
    lbTime.Caption = Format(Hour & ":" & Minute & ":" & Second, "HH:MM:SS AM/PM")
    
    'Set rsL = cnMSSQL.Execute("insert into attlog (acnumber, logdt, readerip)values ('" & EnrollNumber & "', '" & Xdt & "', '" & strIP1 & "')")
    'strFname = EnrollNumber & "_" & Year & "_" & Month & "_" & Day & "_" & Hour & "_" & Minute & "_" & Second & ".jpeg"
    If Not ConnectSQL Then
        If LStudNo.Caption = "" Then
        Else
            LMove
        End If
        lblSql.Caption = "No Database Connection."
        lblSql.ForeColor = vbRed
        lblSql.Visible = True
        LStudNo.Caption = "---"
        LLName.Caption = "---"
        LYear.Caption = "---"
        LNote.Caption = "AC Number - " & lngEnrollNumber & ".  Not Encoded."
        LImage.Picture = LoadPicture(App.Path & "\Users\nophoto.jpg")
    Else
        If LStudNo.Caption = "" Then
        Else
            LMove
        End If
        lblSql.Caption = "Connected to Database."
        lblSql.ForeColor = vbGreen
        lblSql.Visible = False
        PostEntrance EnrollNumber, Xdt
    End If
    lngDiffPIC1 = 0
    dtSRem1 = Now
    
End Sub
Sub PostEntrance(lngEnrollNumber As Long, strXDT As String)
    Dim FSO As New FileSystemObject
    
    Dim rsA As New ADODB.Recordset
    
    Dim rsR As New ADODB.Recordset
    Set rsR = cnMSSQL.Execute("select SSN, Name, Title from userinfo where Badgenumber = '" & lngEnrollNumber & "'")
    If Not rsR.EOF Then
        LStudNo.Caption = rsR!ssn
        LLName.Caption = rsR!Name
        LYear.Caption = IIf(IsNull(rsR!Title), "", rsR!Title)
        'LNote.Caption = rsG!Notes
        Set FSO = New FileSystemObject
        If FSO.FileExists(App.Path & "\Users\" & rsR!ssn & ".jpg") Then
            LImage.Picture = LoadPicture(App.Path & "\Users\" & rsR!ssn & ".jpg")
        Else
            LImage.Picture = LoadPicture(App.Path & "\Users\nophoto.jpg")
        End If
        Set rsA = cnMSSQL.Execute("select * from remarks where ssn='" & rsR!ssn & "' and dtfrom <= '" & strXDT & "' and dtto >= '" & strXDT & "'")
        If Not rsA.EOF Then
            LNote.Caption = rsA!remarks
        Else
            LNote.Caption = vbNullString
        End If
    Else
        LStudNo.Caption = "---"
        LLName.Caption = "---"
        LYear.Caption = "---"
        LNote.Caption = "AC Number - " & lngEnrollNumber & ".  Not Encoded."
        LImage.Picture = LoadPicture(App.Path & "\Users\nophoto.jpg")
    End If
        
        lngDiffPIC1 = 0
        dtSRem1 = Now



End Sub
Sub PostExit(lngEnrollNumber, strXDT As String)
    Dim FSO As New FileSystemObject
    
    'Set FSO = New FileSystemObject
    'If FSO.FileExists(App.Path & "\Users\" & lngEnrollNumber & ".jpg") Then
    '    RImage.Picture = LoadPicture(App.Path & "\Users\" & lngEnrollNumber & ".jpg")
    'Else
    '    RImage.Picture = LoadPicture(App.Path & "\Users\nophoto.jpg")
    'End If
   
    Dim rsG As New ADODB.Recordset
    Set rsG = cnMSSQL.Execute("select SSN, Name, Title from userinfo where Badgenumber = '" & lngEnrollNumber & "'")
    If Not rsG.EOF Then
        RStudNo.Caption = rsG!ssn
        RLName.Caption = rsG!Name
        RYear.Caption = IIf(IsNull(rsG!Title), "", rsG!Title)
        'RNote.Caption = rsG!Notes
        Set FSO = New FileSystemObject
        If FSO.FileExists(App.Path & "\Users\" & rsG!ssn & ".jpg") Then
            RImage.Picture = LoadPicture(App.Path & "\Users\" & rsG!ssn & ".jpg")
        Else
            RImage.Picture = LoadPicture(App.Path & "\Users\nophoto.jpg")
        End If
        Set rsA = cnMSSQL.Execute("select * from remarks where ssn='" & rsG!ssn & "' and dtfrom <= '" & strXDT & "' and dtto >= '" & strXDT & "'")
        If Not rsA.EOF Then
            RNote.Caption = rsA!remarks
        Else
            RNote.Caption = vbNullString
        End If
    Else
        RStudNo.Caption = "---"
        RLName.Caption = "---"
        RYear.Caption = "---"
        RNote.Caption = "AC Number - " & lngEnrollNumber & ".  Not Encoded."
        RImage.Picture = LoadPicture(App.Path & "\Users\nophoto.jpg")
    End If
End Sub

Private Sub ZKExit_OnAttTransaction(ByVal EnrollNumber As Long, ByVal IsInValid As Long, ByVal AttState As Long, ByVal VerifyMethod As Long, ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long)
    Dim Xdt As String
    Dim strFname As String
    Dim rsL As New ADODB.Recordset
    
    Xdt = Format(Month & "/" & Day & "/" & Year & "  " & Hour & ":" & Minute & ":" & Second, "mm/dd/yyyy HH:MM:SS AM/PM")
'    Xdt = Format(Hour & ":" & Minute & ":" & Second, "HH:MM:SS AM/PM")

    'Set rsL = cnMSSQL.Execute("insert into attlog (acnumber, logdt, readerip)values ('" & EnrollNumber & "', '" & Xdt & "', '" & strIP1 & "')")
    'strFname = EnrollNumber & "_" & Year & "_" & Month & "_" & Day & "_" & Hour & "_" & Minute & "_" & Second & ".jpeg"
    rbTime.Caption = Format(Hour & ":" & Minute & ":" & Second, "HH:MM:SS AM/PM")
    
    If Not ConnectSQL Then
        If RStudNo.Caption = "" Then
        Else
            RMove
        End If
        lblSql.Caption = "No Database Connection."
        lblSql.ForeColor = vbRed
        lblSql.Visible = True
        RImage.Picture = LoadPicture(App.Path & "\Users\nophoto.jpg")
        RStudNo.Caption = "---"
        RLName.Caption = "---"
        RYear.Caption = "---"
        RNote.Caption = "AC Number - " & lngEnrollNumber & ".  Not Encoded."
    Else
        If RStudNo.Caption = "" Then
        Else
            RMove
        End If
        lblSql.Caption = "Connected to Database."
        lblSql.ForeColor = vbGreen
        lblSql.Visible = False
        PostExit EnrollNumber, Xdt
    End If
        lngDiffPIC2 = 0
        dtSRem2 = Now
    
    

End Sub
Sub LMove()
    LI2.Picture = LI1.Picture
    LI1.Picture = LImage.Picture
    
    LP2.Caption = LP1.Caption
    LP1.Caption = LStudNo.Caption & " @" & lbTime.Caption & vbCrLf & LLName.Caption



End Sub
Sub RMove()
    RI2.Picture = RI1.Picture
    RI1.Picture = RImage.Picture
    
    RP2.Caption = RP1.Caption
    RP1.Caption = RStudNo.Caption & " @" & rbTime.Caption & vbCrLf & RLName.Caption



End Sub

