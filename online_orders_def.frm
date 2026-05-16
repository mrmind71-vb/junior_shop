VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form orders_online_items 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«’‰«› ÿ·»Ì« "
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   18255
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Index           =   4
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1440
      Width           =   1275
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   45
         TabIndex        =   31
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   820
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_orders_def.frx":0000
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_def.frx":2925
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   465
         Left            =   45
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   630
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   820
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_orders_def.frx":5179
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_def.frx":72D9
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   4770
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   405
      Width           =   6585
      Begin VB.TextBox xNote_branch 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   5145
      End
      Begin VB.TextBox xNote_main 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   5145
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·ÕÊŸ… «·›—⁄"
         Height          =   330
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·ÕÊŸ… «·«œ«—…"
         Height          =   330
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   4770
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1485
      Width           =   6585
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   5190
      End
      Begin VB.TextBox xDelOrder_Date 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "”»» «·«·€«¡"
         Height          =   330
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   675
         Width           =   960
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·«·€«¡"
         Height          =   330
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   315
         Width           =   1005
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   4230
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   8325
      Width           =   2130
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   11385
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   45
      Width           =   6765
      Begin VB.TextBox xphone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   5190
      End
      Begin VB.TextBox xcity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   5190
      End
      Begin VB.TextBox xship_date 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1260
         Width           =   1590
      End
      Begin VB.TextBox xAddress1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   135
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1620
         Width           =   5235
      End
      Begin VB.TextBox xName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   5190
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«· ·Ì›Ê‰"
         Height          =   330
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   945
         Width           =   645
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„œÌ‰…"
         Height          =   330
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·‘Õ‰"
         Height          =   330
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1305
         Width           =   960
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·⁄‰Ê«‰"
         Height          =   330
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1665
         Width           =   645
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·«”„"
         Height          =   330
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   645
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   465
      Left            =   -1845
      Top             =   90
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4740
      Left            =   45
      TabIndex        =   9
      Top             =   2610
      Width           =   18105
      _cx             =   31935
      _cy             =   8361
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   12648384
      ForeColorFixed  =   0
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   8805
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   45
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   476
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   3240
         TabIndex        =   12
         Top             =   45
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   7110
         TabIndex        =   13
         Top             =   45
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCommand CMD_SEND 
      Height          =   510
      Left            =   1350
      TabIndex        =   14
      Top             =   8235
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   900
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "online_orders_def.frx":95C6
      Caption         =   " —ÕÌ· ·›« Ê—… „»Ì⁄« "
      ButtonStyle     =   3
      PictureAlignment=   10
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "online_orders_def.frx":BCE9
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   510
      Left            =   45
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8235
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   900
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "online_orders_def.frx":DFDF
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdNewInv 
      Height          =   510
      Left            =   1350
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8235
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   900
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "online_orders_def.frx":10302
      Caption         =   "«÷«›… „— Ã⁄ «Ê «” »œ«·"
      ButtonStyle     =   3
      PictureAlignment=   9
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "online_orders_def.frx":1230A
   End
   Begin VB.Label xSales_doc 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   540
      Width           =   2130
   End
End
Attribute VB_Name = "orders_online_items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
Public myForm As Form
Dim bEditRecord As Boolean, bAct As Boolean
Dim con As New ADODB.Connection
Dim formMode
Const LoadMode = 0, DefineMode = 1

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdInform_Click()
CardLookup cFilter
End Sub
Private Sub CmdFirst_Click()
If Not openCardTable(tbMode.tbFirst) Then
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
If Not openCardTable Then
    myDefine
End If
End Sub
Private Sub CmdNewInv_Click()
myDefine
End Sub

Private Sub Command1_Click()
   Case 2
        con.Execute " UPDATE FILE6_90H SET [NAME] = " & addstring(.TextMatrix(Row, col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
    Case 3
        con.Execute " UPDATE FILE6_90H SET [PHONE] = " & addstring(.TextMatrix(Row, col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
    Case 4
        con.Execute " UPDATE FILE6_90H SET [CITY] = " & addstring(.TextMatrix(Row, col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
    Case 5
        con.Execute " UPDATE FILE6_90H SET [ADDRESS1] = " & addstring(.TextMatrix(Row, col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
    Case 19
        con.Execute " UPDATE FILE6_90H SET [SHIP_DATE] = " & Add
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then KeyAscii = 0
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End If
End Sub
Private Sub Form_Load()
myLoad
myLoadGrd
End Sub
Private Sub myLoadGrd()
On Error GoTo myerror
Dim storeTable As New ADODB.Recordset
Dim cString As String
If cBranch = "00" Then
    If myForm.Check1.Value = 0 Then
        cString = " SELECT *" & _
                  " FROM FILE0_40 " & _
                  " WHERE ONLINE = 1" & _
                  " ORDER BY CODE "
    Else
        cString = "SELECT FILE0_40.CODE," & _
                  "FILE0_40.DESCA, " & _
                  "FILE0_40.[ONLINE]" & _
                  " FROM FILE0_40" & _
                  " WHERE FILE0_40.CODE IN (" & _
                  " SELECT FILE1_11.STORE " & _
                  " FROM FILE1_11 " & _
                  " INNER JOIN FILE6_90 ON FILE6_90.ITEM = FILE1_11.ITEM " & _
                  " WHERE FILE6_90.DOC_NO = " & MyParn(sDoc_no) & _
                  " GROUP BY FILE1_11.STORE," & _
                  " FILE1_11.ITEM " & _
                  " HAVING SUM([IN] - [OUT]) <> 0)" & _
                  " AND ISONEST = 0 " & _
                  " AND ISONEST2 = 0" & _
                  " AND ISSTOP = 0" & _
                  " AND FILE0_40.BRANCH IS NOT NULL" & _
                  " AND FILE0_40.CODE <> '000'" & _
                  " OR FILE0_40.ONLINE = 1" & _
                  " GROUP BY " & _
                  " FILE0_40.CODE," & _
                  " FILE0_40.DESCA, " & _
                  " FILE0_40.ONLINE" & _
                  " ORDER BY FILE0_40.[ONLINE] DESC," & _
                  " FILE0_40.CODE"
    End If
Else
    cString = " SELECT * " & _
              " FROM FILE0_40 " & _
              " WHERE CODE = " & MyParn(cBranchStore) & _
              " ORDER BY CODE "
End If

Set storeTable = myRs(cString)
Do Until storeTable.EOF
    cField = cField & _
            "(" & _
            " SELECT SUM([IN]-[OUT])" & _
            " FROM FILE1_11" & _
            " WHERE FILE1_11.ITEM = FILE6_90.ITEM" & _
            " AND FILE1_11.STORE  = " & MyParn(storeTable!code) & _
            " ) AS " & addstring(storeTable!DESCA) & ","
            
    storeTable.MoveNext
Loop

With grid1
    cString = " SELECT FILE6_90.ITEM," & _
              " FILE6_90.SKU," & _
              "FACT.DESCA," & _
              " COALESCE(FILE1_10.desca,FILE6_90.ITEM_NAME)," & _
              " FILE1_10.COLOR," & _
              " FILE1_10.SCAL," & _
              " FILE6_90.QUANT," & _
              " FILE6_90.PRICE," & _
              " file6_90.ITEM_NAME," & _
              " FILE6_90.ID, " & _
              " v.QUANT AS [Total Orders], " & _
              cField & _
              "(SELECT SUM(QUANT) FROM Q_BAL_NOTSHIP WHERE Q_BAL_NOTSHIP.ITEM = FILE6_90.ITEM)" & _
              " FROM  FILE6_90 " & _
              " INNER JOIN FILE6_90H ON FILE6_90.DOC_NO = FILE6_90H.DOC_NO " & _
              " LEFT  JOIN FILE1_10 ON FILE6_90.ITEM = FILE1_10.ITEM" & _
              " LEFT  JOIN  FACT ON FILE1_10.code = FACT.CODE" & _
              " LEFT JOIN vw_online_orders_open_quant as v ON FILE6_90.ITEM = V.ITEM AND FILE6_90H.STORE = V.STORE" & _
              " WHERE FILE6_90H.DOC_NO =  " & MyParn(sDoc_no)
    Set grid1.DataSource = myRs(cString)
End With

Fixgrd

storeTable.Close
Set storeTable = Nothing
grid1.Cell(flexcpAlignment, 0, 0, grid1.Rows - 1, grid1.Cols - 1) = 7
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myLoad()
Dim loctable As New ADODB.Recordset
Set loctable = myRs("SELECT FILE6_90H.* " & _
                    " FROM FILE6_90H WHERE DOC_NO = " & MyParn(sDoc_no))
If Not loctable.EOF Then
    xName.text = loctable!Name & ""
    xphone.text = loctable!phone & ""
    xcity.text = loctable!city & ""
    xAddress1.text = loctable!address1 & ""
    xship_date.text = myFormat_p(loctable!ship_date)
    xNotes.text = loctable!notes & ""
    xNote_branch.text = loctable!note_branch & ""
    xNote_main.text = loctable!note_main & ""
    xDelOrder_Date.text = myFormat_p(loctable!delorder_date)
    xSales_doc.Caption = loctable!sales_doc & ""
End If
cmdSave.Enabled = xSales_doc.Caption = ""
loctable.Close
Set loctable = Nothing
End Sub
Sub Fixgrd()
Dim bShowColAll As Boolean, bShowCol As Boolean
With grid1
        
    .RowHeight(0) = 800
    .WordWrap = True
    .RowHidden(1) = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "SKU"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "«·’‰›"
    .TextMatrix(0, 4) = "«··Ê‰"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«·ﬂ„Ì…"
    .TextMatrix(0, 7) = "«·”⁄—"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 4000
    .ColWidth(4) = 1500
    .ColWidth(5) = 800
    .ColWidth(6) = 600
    .ColWidth(7) = 700
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 1000
    
    .ColWidth(.Cols - 1) = 1000
    
    .TextMatrix(0, .Cols - 1) = "—’Ìœ ‘Õ‰ €Ì— „” ·„"
    
    nColWidth = IIf(cBranch = "00", 1000, 1000)
        
    For col = 11 To .Cols - 1
        .ColWidth(col) = nColWidth
    Next
        
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .ColDataType(0) = flexDTDate
    .SubtotalPosition = flexSTBelow
    
    .Subtotal flexSTSum, -1, 6, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 8, "#", vbYellow, , True, ""
    
    .Select .Rows - 1, 0
    .ShowCell .Rows - 1, 0
End With
End Sub

Private Sub SSCommand1_Click()

End Sub
