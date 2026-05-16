VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form online_orders_sub 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«’‰«› ÿ·»Ì« "
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   18255
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   0
      Width           =   3705
      Begin Threed.SSCommand CMD_SEND 
         Height          =   510
         Left            =   1350
         TabIndex        =   18
         Top             =   135
         Width           =   2310
         _ExtentX        =   4075
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
         Picture         =   "online_orders_invoices.frx":0000
         Caption         =   " —ÕÌ· ·›« Ê—… „»Ì⁄« "
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":2723
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   135
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
         Picture         =   "online_orders_invoices.frx":4A19
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   6345
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   510
         Left            =   2700
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   900
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
         Picture         =   "online_orders_invoices.frx":6D3C
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_orders_invoices.frx":8EE3
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   510
         Left            =   1800
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   900
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
         Picture         =   "online_orders_invoices.frx":AF2A
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_orders_invoices.frx":D015
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   510
         Left            =   945
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   900
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
         Picture         =   "online_orders_invoices.frx":F00F
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_orders_invoices.frx":11120
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   510
         Left            =   45
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   900
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
         Picture         =   "online_orders_invoices.frx":1311A
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_orders_invoices.frx":1533E
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   12915
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   5325
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   2700
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   135
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
         Picture         =   "online_orders_invoices.frx":1740F
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":19417
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1350
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
         _ExtentX        =   2328
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
         Picture         =   "online_orders_invoices.frx":1B3CE
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":1DB6A
      End
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4005
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
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
         Picture         =   "online_orders_invoices.frx":1FFFE
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":223C9
      End
   End
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Index           =   4
      Left            =   9900
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   3030
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   1530
         TabIndex        =   6
         Top             =   135
         Width           =   1455
         _ExtentX        =   2566
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
         Picture         =   "online_orders_invoices.frx":24472
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":26D97
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   135
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   900
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
         Picture         =   "online_orders_invoices.frx":295EB
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":2B74B
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
      Height          =   6855
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   18150
      _cx             =   32015
      _cy             =   12091
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
      ScrollBars      =   2
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
      TabIndex        =   1
      Top             =   7590
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
End
Attribute VB_Name = "online_orders_sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nType As Byte, bedit As Boolean, sDoc_no As String
Public bCheck As Boolean
Dim bEditRecord As Boolean, bAct As Boolean
Dim oSeachCode As New Search, oSearchDoc As New Search, oSearchProject As New Search, oSearchClient As New Search
Dim cFilter As String, cSelect As String
Dim con As New ADODB.Connection, CardTable As ADODB.Recordset
Dim cFile As String, cFileHeader As String, aBoxs As Variant
Dim clist1 As String, cList2 As String, cList3 As String, sDef As String
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, IIf(XDOC_NO.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cUserName))
aInsert = AddFlag(aInsert, IIf(XDOC_NO.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")
con.BeginTrans
On Error GoTo myerror
If XDOC_NO.Tag = DefineMode Then
    XDOC_NO.text = Newflag(cFileHeader, "DOC_NO", con)
    aInsert = AddFlag(aInsert, "DOC_NO", addvalue(XDOC_NO.text))
    con.Execute addInsert(aInsert, cFileHeader)
Else
    con.Execute addUpdate(aInsert, cFileHeader, "doc_no = " & addvalue(XDOC_NO.text))
End If
myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc(Optional sControl As String = "")
If ActiveControl.Name = grid1.Name Then
    grid1.TextMatrix(grid1.Row, 2) = oSeachCode.grid1.TextMatrix(oSeachCode.grid1.Row, 0)
    Grid1_AfterEdit grid1.Row, grid1.Col
    oSeachCode.Hide
    CellPos 13, grid1.Row, grid1.Col
ElseIf ActiveControl.Name = cmdInform.Name Then
    openCardTable tbMode.tbFind, oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
ElseIf ActiveControl.Name = cmdClient.Name Then
    If oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0) = "" Then
        cmdClient.Tag = ""
        cmdClient.Caption = cmdClient.TagVariant
    Else
        cmdClient.Tag = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
        cmdClient.Caption = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 1)
    End If
    If Not openCardTable(tbMode.tbFind, XDOC_NO.text) Then
        If Not openCardTable Then myDefine
    End If
    oSearchClient.Hide
ElseIf ActiveControl.Name = cmdProject.Name Then
    If oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 0) = "" Then
        cmdProject.Tag = ""
        cmdProject.Caption = cmdProject.TagVariant
    Else
        cmdProject.Tag = oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 0)
        cmdProject.Caption = oSearchProject.grid1.TextMatrix(oSearchProject.grid1.Row, 1)
    End If
    If Not openCardTable(tbMode.tbFind, XDOC_NO.text) Then
        If Not openCardTable Then myDefine
    End If
    oSearchProject.Hide
End If
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› ?", vbOKCancel + vbDefaultButton2 + vbCritical) <> vbOK Then Exit Sub

con.BeginTrans
On Error GoTo myerror
con.Execute "Delete  From " & cFile & " where Doc_No = " & addvalue(XDOC_NO.text)
con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & addvalue(XDOC_NO.text)
con.CommitTrans

If sDoc_no <> "" Then Exit Sub

If Not openCardTable(tbMode.tbPrevious, XDOC_NO.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
cString = "SELECT Doc_No, Convert(Varchar(10),DATE,111), [dbo].[fn_client_names_cash](DOC_NO," & nType & ")" & _
          " FROM " & cFileHeader
If pFilter <> "" Then cString = cString & " WHERE " & pFilter
Generalarray(1) = cString
Generalarray(2) = "order by DATE ,DOC_NO "
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì·"
listarray(0, 1) = "DOC_NO IN (SELECT DOC_NO FROM " & cFile & " INNER JOIN FILE3_10 ON " & cFile & ".CODE = FILE3_10.CODE WHERE %%FILE3_10.DESCA%%)"

listarray(1, 0) = " «—ÌŒ «·„” ‰œ"
listarray(1, 1) = "(##[Date]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1350

GrdArray(2, 0) = "«·≈”„"
GrdArray(2, 1) = 9000

Dim aFilter As Variant
aFilter = AddFlag(aFilter, "FILTER", True)
aFilter = AddFlag(aFilter, "FIELD", "DOC_NO")

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.sCaption = IIf(nType = 0, "«” ⁄·«„ „ﬁ»Ê÷«  ‰ﬁœÌ… „‰ ⁄„·«¡", "«” ⁄·«„ „œ›Ê⁄«  ‰ﬁœÌ… ≈·Ì «·⁄„·«¡")
oSearchDoc.aFilter = aFilter
oSearchDoc.Show 1
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
Private Sub CmdNext_Click()
If Not openCardTable(tbMode.tbNext, XDOC_NO.text) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, XDOC_NO.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdNewInv_Click()
myDefine
End Sub
Private Sub cmdProject_Click()
ProjectLookup Me, oSearchProject, , , cmdProject.Tag <> ""
End Sub
Private Sub cmdSave_Click()
If Not myValid Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If sDoc_no <> "" Then
    Unload Me
    Exit Sub
End If
If Not openCardTable(tbMode.tbFind, XDOC_NO.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub cmdClient_Click()
CLIENTLOOKUP Me, oSearchClient, , , cmdClient.Tag <> ""
End Sub

Private Sub CmdUndo_Click()
'openCardTable
myUndo
End Sub

Private Sub Command1_Click()
End Sub
Private Sub cmdYear_Click()
Set oSearchYear.myForm = Me
oSearchYear.bEmpty = cmdYear.Tag <> ""
oSearchYear.Show 1
End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If XDOC_NO.Tag = LoadMode Then
        grid1.SetFocus
    Else
        xDate.SetFocus
    End If
    Err.Clear
End If
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
bedit = True
cFile = IIf(nType = 0, "FILE8_10", "FILE8_30")
cFileHeader = IIf(nType = 0, "FILE8_10H", "FILE8_30H")
openCon con

clist1 = StrList2("select code,desca from Projects order by code")
cList2 = StrList2("Select code,desca from file0_50")
cList3 = StrList2("select code,desca from FILE3_10 order by desca")

Set grid1.DataSource = data1

If Not openCardTable Then myDefine

Me.Caption = IIf(nType = 0, "„œ›Ê⁄«  «·Ì «·⁄„·«¡", "„ﬁ»Ê÷«  „‰ «·⁄„·«¡")
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveText Me, , Array(chkYear.Name, chkMonth.Name, chkYear.Name, chkBalance.Name, cmdYear.Name)
CardTable.Close
Set CardTable = Nothing
closeCon con
Set cash_supfrm = Nothing
Err.Clear
End Sub
Private Function myValid(Optional bIgMsg As Boolean = True) As Boolean
If Trim(XDOC_NO.text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

myValid = True
End Function
Private Sub myload()
XDOC_NO.text = CardTable!Doc_no
xDate.text = myFormat_p(CardTable!Date)
panel1(0).Caption = CardTable!UserName & " " & myFormat_p(CardTable!Time, True)
panel1(1).Caption = CardTable!UserName2 & " " & myFormat_p(CardTable!time2, True)
Handlecontrols LoadMode
myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
XDOC_NO.text = Newflag(cFileHeader, "DOC_NO", con)
xDate.text = myFormat_p(Date)
Fixgrd

grid1.Rows = 1

myAddItem

Handlecontrols DefineMode
'StatusBar1.Panels(2).text = "·« ÌÊÃœ ”œ«œ"
'StatusBar1.Panels(3).text = ""
'StatusBar1.Panels(4).text = ""

CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bedit
cmdFilter.Visible = cmdFilter.Tag <> ""
cmdNewInv.Enabled = nMode = LoadMode And bedit
cmddel.Enabled = nMode = LoadMode And bEditRecord

cmdSave.Enabled = bEditRecord
cmddel.Enabled = nMode = LoadMode And bEditRecord

Dim nRecord As Long, nRecords As Long
retRecords XDOC_NO.text, nRecords, nRecord

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
     panel1(0).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""

XDOC_NO.Enabled = (nMode = DefineMode)
XDOC_NO.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus XDOC_NO
If Not IsDgt(XDOC_NO.text) Then
     If XDOC_NO.Tag = LoadMode Then
        myDefine
    Else
        XDOC_NO.text = ""
    End If
Else
    If (Not (CardTable.EOF)) And XDOC_NO.Tag = LoadMode Then
        If CardTable!Doc_no = XDOC_NO.text Then
            Exit Sub
        End If
    End If
    
    openCardTable XDOC_NO.text
    If Not CardTable.EOF Then
        myload
    ElseIf XDOC_NO.Tag = LoadMode Then
        myDefine
    Else
'        xDoc_No.Text = ""
    End If
End If
End Sub
Private Function CalcTotals(Optional Row As Long = 0)
Dim nTotal As Double
With grid1
Dim I As Long
For I = 1 To .Rows - 2
    nTotal = nTotal + mRound(.TextMatrix(I, 3))
Next

If Row = 0 Then Exit Function

.TextMatrix(Row, 6) = ""
.TextMatrix(Row, 7) = ""

If .TextMatrix(Row, 2) = "" Then Exit Function
.TextMatrix(Row, 6) = MyFuncValue("dbo.fn_client_balance", con, .TextMatrix(Row, 2), "NULL", "NULL")
If .TextMatrix(Row, 0) <> "" Then
    .TextMatrix(Row, 7) = MyFuncValue("dbo.fn_client_balance", con, .TextMatrix(Row, 2), .TextMatrix(Row, 0), "NULL")
End If
End With
End Function
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_no As String = "", Optional bDefine As Boolean = True) As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_no) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

On Error GoTo myerror
cFilter = ""
If cmdFilter.Tag <> "" Then cFilter = cFilter & Tr(cFilter) & "DOC_NO IN (" & cmdFilter.Tag & ")"
If cmdClient.Tag <> "" Then cFilter = cFilter & Tr(cFilter) & "DOC_NO IN (SELECT " & cFile & ".DOC_NO FROM " & cFile & " WHERE " & cFile & ".CODE = " & cmdClient.Tag & ")"
If cmdProject.Tag <> "" Then cFilter = cFilter & turn(cFilter, " And ") & "DOC_NO IN (SELECT " & cFile & " FROM " & cFile & " WHERE " & cFile & ".PROJECT = " & cmdProject.Tag & ")"
If xYear.text <> "" Then cFilter = cFilter & Tr(cFilter) & "YEAR(DATE) = " & xYear.text
If xMonth.text <> "" Then cFilter = cFilter & Tr(cFilter) & "MONTH(DATE) = " & xMonth.text

If sDoc_no <> "" Then cFilter = "DOC_NO = " & sDoc_no
cString = "SELECT TOP 1 * " & _
          " FROM " & cFileHeader

If pMode = tbMode.tbFirst Then
    cOrder = "Order by Doc_no"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by Doc_no DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "Doc_no = " & pDoc_no
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "Doc_no < " & pDoc_no
    cOrder = "Order by Doc_no desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "Doc_no > " & pDoc_no
    cOrder = "Order by Doc_no"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

If cFilter <> "" Then
    cString = cString & IIf(cWhere = "", " WHERE ", " AND ") & cFilter
End If

cString = cString & " " & cOrder

Set CardTable = mycmd(cString, con, adText)

If (Not CardTable.EOF) Then
    myload
    openCardTable = True
End If
Me.MousePointer = 0
Exit Function
myerror:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function
Private Sub myUndo()
If XDOC_NO.Tag = DefineMode Then
    If Not openCardTable Then
        CmdNewInv_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, XDOC_NO.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus XDOC_NO
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xDate_DblClick()
Set datefrm.oDate = xDate
datefrm.Show 1
End Sub
Private Sub chkDay_Click()
If Not openCardTable(tbMode.tbFind, XDOC_NO.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub chkMonth_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, XDOC_NO.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Sub chkOpen_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, XDOC_NO.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Sub chkYear_Click()
If Not bCheck Then
    If Not openCardTable(tbMode.tbFind, XDOC_NO.text) Then
        If Not openCardTable Then myDefine
    End If
End If
End Sub
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_no <> "" Then
    cString = "SELECT Count(*) AS records,SUM(CASE WHEN DOC_NO <= " & pDoc_no & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM " & cFileHeader & Tr(cFilter, " WHERE ") & cFilter

Set loctable = mycmd(cString, con)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Sub cmdFilter_Click()
cmdFilter.Tag = ""
If Not openCardTable(tbMode.tbFind, XDOC_NO.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Sub myproc2(pFilter As String)
oSearchDoc.Hide
cmdFilter.Tag = pFilter
If Not openCardTable(tbMode.tbFirst, XDOC_NO.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub myreplaceGrd(Row As Long)
Dim aInsert As Variant
With grid1
    For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(XDOC_NO.text))
        aInsert = AddFlag(aInsert, "Project", addvalue(grid1.TextMatrix(I, 0)))
        aInsert = AddFlag(aInsert, "Box", addstring(grid1.TextMatrix(I, 1)))
        aInsert = AddFlag(aInsert, "Code", addstring(grid1.TextMatrix(I, 2)))
        aInsert = AddFlag(aInsert, "[VALUE]", mRound(grid1.TextMatrix(I, 3)))
        aInsert = AddFlag(aInsert, "Receipt", addstring(grid1.TextMatrix(I, 4)))
        aInsert = AddFlag(aInsert, "Desca", addstring(grid1.TextMatrix(I, 5)))
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, cFile)
        Else
            con.Execute addUpdate(aInsert, cFile, "ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not myValid(True) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myLoadGrd
    If Row < grid1.Rows - 1 Then
        grid1.Select Row, Col
    Else
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If


If Not validRow(Row, Col) Then
    CalcTotals Row
    Exit Sub
End If

With grid1
If Row = grid1.Rows - 1 Then
    myAddItem
ElseIf Row = grid1.Rows - 2 And (Col = 0 Or Col = 1) Then
    MyEditItem grid1, Row, Col
End If

If myreplace(Row) Then
    If XDOC_NO.Tag = DefineMode Then
        Handlecontrols LoadMode
        myLoadGrd
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myLoadGrd
    Else
        CalcTotals Row
    End If
Else
    myLoadGrd
End If
End With
End Sub
Private Sub Grid1_EnterCell()
If Not bEditRecord Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub myLoadGrd()
Dim cString As String
Dim aFields(8)
aFields(0) = "PROJECT"
aFields(1) = "[BOX]"
aFields(2) = "CODE"
aFields(3) = "[VALUE]"
aFields(4) = "RECEIPT"
aFields(5) = "DESCA"
aFields(6) = "dbo.fn_client_balance(CODE,NULL,NULL)"
aFields(7) = "dbo.fn_client_balance(CODE,PROJECT,NULL)"
aFields(8) = "ID"

cString = "Select " & arString(aFields, ",") & _
         " FROM " & cFile
cString = cString & " WHERE  DOC_NO = " & addvalue(XDOC_NO.text)
Set data1.Recordset = mycmd(cString, con)
myAddItem

CalcTotals
Fixgrd
End Sub
Private Sub Grid1_GotFocus()
Grid1_EnterCell
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 112 And grid1.Col = 2 Then
    ClientLookupAll Me, oSeachCode
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from " & cFile & " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        myRemove grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
        MsgBox "«·„‘—Ê⁄ €Ì— „”Ã·"
    End If
ElseIf Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
        MsgBox "«·Œ“‰… €Ì— „”Ã·…"
    End If
End If
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "«·„‘—Ê⁄|" & "Œ“‰…|" & "«·⁄„Ì·|" & "«·ﬁÌ„…|" & "—ﬁ„ «·«Ì’«·|" & "«·»Ì«‰|" & "—’Ìœ «·„‘—Ê⁄|" & "—’Ìœ «·⁄„Ì·|"
.ColWidth(0) = 2500
.ColWidth(1) = 2000
.ColWidth(2) = 2500
.ColWidth(3) = 1200
.ColWidth(4) = 1600
.ColWidth(5) = 4000
.ColWidth(6) = 1400
.ColWidth(7) = 1400
.ColHidden(.Cols - 1) = True
For I = 1 To grid1.Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
.ColComboList(0) = clist1
.ColComboList(1) = cList2
.ColComboList(2) = cList3
End With
End Sub
Private Sub myAddItem()
With grid1
.AddItem ""
If grid1.Rows > 2 Then
    .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
    .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 2, 1)
Else
    .TextMatrix(.Rows - 1, 0) = "1"
    .TextMatrix(.Rows - 1, 1) = "1"
End If
End With
End Sub
Private Function validRow(Row As Long, Optional Col As Long = -1) As Boolean
With grid1
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 2)) = "" Then Exit Function
If mRound(.TextMatrix(Row, 3)) = 0 Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 4 Then
    grid1.Col = Col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 3)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
CalcTotals
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If Col = 0 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    If Col = 1 And (grid1.TextMatrix(Row, Col) = "" Or grid1.TextMatrix(Row, grid1.Cols - 1) = "") Then Exit Sub
    If Col = 2 And (grid1.TextMatrix(Row, Col) = "" Or grid1.TextMatrix(Row, grid1.Cols - 1) = "") Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid1.Col = 0 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    If grid1.Col = 1 And (grid1.TextMatrix(grid1.Row, grid1.Col) = "" Or grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Then Exit Sub
    If grid1.Col = 2 And (grid1.TextMatrix(grid1.Row, grid1.Col) = "" Or grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Then Exit Sub
    KeyAscii = 0
End If
End Sub
