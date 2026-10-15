VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdSalesReturnfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "„—«Ã⁄… »Ê‰«  «·„— Ã⁄«  ··›—Ê⁄"
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   675
      Width           =   4605
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   135
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   979
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
         Picture         =   "grdSalesReturn.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1125
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   135
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   979
         _Version        =   196610
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
         Picture         =   "grdSalesReturn.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdSalesReturn.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   3375
         TabIndex        =   20
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   979
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
         Picture         =   "grdSalesReturn.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2205
         TabIndex        =   21
         Top             =   135
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   979
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
         Picture         =   "grdSalesReturn.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1410
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11265
      Begin VB.TextBox XDATER2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   1500
      End
      Begin VB.TextBox XDATER1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   1680
      End
      Begin VB.TextBox XDOC_NO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   3300
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Visible         =   0   'False
         Width           =   1590
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   7155
         TabIndex        =   8
         Top             =   945
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XSTORERET 
         Height          =   315
         Left            =   225
         TabIndex        =   10
         Top             =   945
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XUSER 
         Height          =   315
         Left            =   225
         TabIndex        =   14
         Top             =   585
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·’·«ÕÌ…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "›—⁄ «·„— Ã⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   990
         Width           =   930
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "›—⁄ «·„»Ì⁄« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9855
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   990
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ ›« Ê—… »Ì⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9855
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… ··„— Ã⁄ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… ··„»Ì⁄« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9855
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6495
      Left            =   90
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1440
      Width           =   20220
      _cx             =   35666
      _cy             =   11456
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
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
End
Attribute VB_Name = "grdSalesReturnfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String
    cHead1 = "„—«Ã⁄… »Ê‰«  «·„— Ã⁄ " & xStore.text
    cHead2 = " »Ì⁄ „‰  «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY") & " Õ Ï  «—ÌŒ " & Format(xdate2.text, "DD-MM-YYYY")
    cHead3 = " „— Ã⁄ „‰  «—ÌŒ " & Format(XDATER1.text, "DD-MM-YYYY") & " Õ Ï  «—ÌŒ " & Format(XDATER2.text, "DD-MM-YYYY")
    
    Load PrintGrd
    PrintGrd.doprint grid1, 1, , cHead1, cHead2, cHead3, , True, 10
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    OpenCon con
    
    Set grid1.DataSource = DATA6
    DATA6.ConnectionString = strCon
    
    data1.ConnectionString = strCon
    data1.RecordSource = "SELECT * FROM STORE_BR ORDER BY CODE "
    Set xStore.RowSource = data1
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    If cBranch <> "00" Then
        xStore.BoundText = cBranchStore
        xStore.Enabled = True
    End If
    
    Set XSTORERET.RowSource = data1
    XSTORERET.ListField = "Desca"
    XSTORERET.BoundColumn = "Code"
    If cBranch <> "00" Then
        XSTORERET.BoundText = cBranchStore
        XSTORERET.Enabled = False
    End If
 
    DATA10.ConnectionString = strCon
    DATA10.RecordSource = "SELECT username_ret FROM QFILE6_20H GROUP BY username_ret ORDER BY username_ret  "
    Set XUSER.RowSource = DATA10
    XUSER.ListField = "username_ret"
    XUSER.BoundColumn = "username_ret"
 
    grid1.Rows = 1
    grid1.Cols = 9
    FIXGRID
End Sub
Private Sub myload()
'On Error GoTo myerror
Dim i As Double
Dim cString  As String, cStr2 As String
With grid1
'                           0               1                       2                3
    
    cString = " SELECT        store, STORE_BR.DESCA, [date], DOC_NO , (SELECT SUM(QUANT) FROM QFILE6_20 WHERE DOC_NO = QFILE6_20H_RET.DOC_NO  )  , CASH + VISA AS TOTAL, CASH,VISA , PHONE , username_ret , SALES_RET " & _
            " , (SELECT DATE FROM QFILE6_20H WHERE DOC_NO = QFILE6_20H_RET.SALES_RET ) " & _
            " , (SELECT STORE FROM QFILE6_20H WHERE DOC_NO = QFILE6_20H_RET.SALES_RET ) AS STORE_RET " & _
            " , (SELECT SUM(QUANT) FROM QFILE6_20 WHERE DOC_NO = QFILE6_20H_RET.SALES_RET ) " & _
            " , (SELECT CASH+VISA FROM QFILE6_20H WHERE DOC_NO = QFILE6_20H_RET.SALES_RET ) " & _
            " , (SELECT CASH FROM QFILE6_20H WHERE DOC_NO = QFILE6_20H_RET.SALES_RET ) " & _
            " , (SELECT VISA FROM QFILE6_20H WHERE DOC_NO = QFILE6_20H_RET.SALES_RET ) " & _
            " , (SELECT PHONE FROM QFILE6_20H WHERE DOC_NO = QFILE6_20H_RET.SALES_RET ) " & _
             "  FROM            QFILE6_20H_RET INNER JOIN STORE_BR ON QFILE6_20H_RET.store = STORE_BR.CODE WHERE DOC_NO IS NOT NULL"
'    If IsDate(xDate1.Text) Then cString = cString & " AND Q_RET_SALES.DATE >= " & DateSq(xDate1.Text)
'    If IsDate(xdate2.Text) Then cString = cString & " AND Q_RET_SALES.DATE <= " & DateSq(xdate2.Text)
    
    If IsDate(XDATER1.text) Then cString = cString & " AND DATE >= " & DateSq(XDATER1.text)
    If IsDate(XDATER2.text) Then cString = cString & " AND DATE <= " & DateSq(XDATER2.text)
     
'    If xStore.BoundText <> "" Then cString = cString & " AND Q_RET_SALES.STORE = " & MyParn(xStore.BoundText)
    If XSTORERET.BoundText <> "" Then cString = cString & " AND STORE = " & MyParn(XSTORERET.BoundText)
    
    If xDoc_No.text <> "" Then cString = cString & " AND SALES_RET = " & MyParn(xDoc_No.text)
    If XUSER.text <> "" Then cString = cString & " AND username_ret = " & MyParn(XUSER.text)

    
    cString = cString & " ORDER BY store, date , DOC_NO "
    DATA6.RecordSource = cString
    DATA6.Refresh
End With
FIXGRID
If grid1.Rows > 1 Then grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub FIXGRID()
With grid1
    .RowHeight(0) = 1000
    .WordWrap = True
    .Cols = 18
    
    .TextMatrix(0, 0) = "—ﬁ„"
    .TextMatrix(0, 1) = "«·›—⁄"
    .TextMatrix(0, 2) = " «—ÌŒ"
    .TextMatrix(0, 3) = "„” ‰œ „— Ã⁄"
    .TextMatrix(0, 4) = "⁄œœ „— Ã⁄"
    .TextMatrix(0, 5) = "ﬁÌ„… »Ê‰ «·„— Ã⁄"
    .TextMatrix(0, 6) = "”œ«œ ‰ﬁœÏ"
    .TextMatrix(0, 7) = "›Ì“«"
    .TextMatrix(0, 8) = "  ·Ì›Ê‰"
    .TextMatrix(0, 9) = " «·’·«ÕÌ…"
    
    .TextMatrix(0, 10) = "»Ê‰ „»Ì⁄« "
    .TextMatrix(0, 11) = " «—ÌŒ"
    .TextMatrix(0, 12) = " «·›—⁄"
    .TextMatrix(0, 13) = "⁄œœ "
    .TextMatrix(0, 14) = " «·ﬁÌ„…"
    .TextMatrix(0, 15) = "”œ«œ ‰ﬁœÏ"
    .TextMatrix(0, 16) = "›Ì“«"
    .TextMatrix(0, 17) = "  ·Ì›Ê‰"
    
    .ColWidth(0) = 700
    .ColWidth(1) = 1500
    .ColWidth(2) = 1400
    .ColWidth(3) = 1600
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    .ColWidth(9) = 1500
    .ColWidth(10) = 1600
    .ColWidth(11) = 1300
    .ColWidth(12) = 1600
    .ColWidth(13) = 500
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    .ColWidth(16) = 1000
    .ColWidth(17) = 1500

    
    .MergeCells = flexMergeFree
    .ColComboList(12) = StrList("SELECT CODE , DESCA FROM STORE_BR", con)
    .MergeCol(3) = True
  
    .ColDataType(3) = flexDTDate
    .ColDataType(8) = flexDTDate
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con
End Sub
Private Sub grid1_DblClick()
With grid1
    If .col = 3 Then
        salesfrm.sDoc_no = .TextMatrix(.Row, 3)
        salesfrm.Show
    End If
    If .col = 10 Then
        salesfrm.sDoc_no = .TextMatrix(.Row, 10)
        salesfrm.Show
    
    End If
    
End With
End Sub
