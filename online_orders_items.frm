VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form orders_online_items 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«’‰«› ÿ·»Ì« "
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18870
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
   ScaleHeight     =   8325
   ScaleWidth      =   18870
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   600
      ScaleWidth      =   18870
      TabIndex        =   34
      Top             =   7350
      Width           =   18870
      Begin Threed.SSCommand cmdSend 
         Height          =   510
         Left            =   90
         TabIndex        =   35
         Top             =   90
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
         Picture         =   "online_orders_items.frx":0000
         Caption         =   " —ÕÌ· ·›« Ê—… „»Ì⁄« "
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_items.frx":2723
      End
   End
   Begin VB.Frame fmAddInvoice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   585
      Width           =   3030
      Begin Threed.SSCommand cmdAddInvoice 
         Height          =   555
         Left            =   90
         TabIndex        =   9
         Top             =   225
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   979
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«÷«›… ÿ·»Ì…"
         ButtonStyle     =   2
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_items.frx":4A19
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1395
      Width           =   1635
      Begin Threed.SSCommand cmdExit 
         Height          =   960
         Left            =   45
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1693
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
         Picture         =   "online_orders_items.frx":6D0F
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
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
      Left            =   1710
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1440
      Width           =   1410
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   45
         TabIndex        =   10
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "online_orders_items.frx":9032
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_items.frx":B957
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   465
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "online_orders_items.frx":E1AB
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_items.frx":1030B
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1500
      Left            =   3195
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   45
      Width           =   6585
      Begin VB.TextBox xNote_branch 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   540
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   855
         Width           =   5145
      End
      Begin VB.TextBox xNote_main 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   5145
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·ÕÊŸ… «·›—⁄"
         Height          =   330
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   810
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
      Height          =   1050
      Left            =   3195
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1530
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
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1320
      End
      Begin Threed.SSCommand cmdDelCancel 
         Height          =   330
         Left            =   2340
         TabIndex        =   37
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " —«Ã⁄ ⁄‰ «·«·€«¡"
         TagVariant      =   "«Œ «— ‰Ê⁄ «·ÿ·»Ì…"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         ShapeSize       =   1
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
      Left            =   9810
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   45
      Width           =   8970
      Begin VB.CheckBox xClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„€·ﬁ…"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   585
         Width           =   780
      End
      Begin VB.TextBox xphone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3690
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   4065
      End
      Begin VB.TextBox xcity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3690
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   4065
      End
      Begin VB.TextBox xship_date 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   180
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1590
      End
      Begin VB.TextBox xStreet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   1170
         Left            =   90
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1260
         Width           =   7665
      End
      Begin VB.TextBox xName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3690
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   4065
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«· ·Ì›Ê‰"
         Height          =   330
         Left            =   7830
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   945
         Width           =   645
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„œÌ‰…"
         Height          =   330
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·‘Õ‰"
         Height          =   330
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   180
         Width           =   960
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·⁄‰Ê«‰"
         Height          =   330
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1305
         Width           =   645
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·«”„"
         Height          =   330
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   180
         Width           =   645
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   465
      Left            =   -3420
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
      Height          =   4650
      Left            =   45
      TabIndex        =   13
      Top             =   2610
      Width           =   18735
      _cx             =   33046
      _cy             =   8202
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
      TabIndex        =   14
      Top             =   7950
      Width           =   18870
      _ExtentX        =   33285
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   17
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
      Left            =   -495
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   180
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "orders_online_items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
Public sStore As String
Public bAllStore As Boolean
Dim bEditRecord As Boolean, bAct As Boolean
Dim oSalesRefund As New sales_refundfrm
Dim con As New ADODB.Connection
Dim bUpdated As Boolean
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Sub cmdAddInvoice_Click()
orders_online_invoices.bEdit = True
orders_online_invoices.sOrder_no = sDoc_no
orders_online_invoices.sStore = sStore
Set orders_online_invoices.myForm = Me
orders_online_invoices.Show 1
End Sub

Private Sub cmdDelCancel_Click()
If MsgBox(" —«Ã⁄ ⁄‰ «·«·€«¡", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
If cancelDelete Then
    Inform " „ «· —«Ã⁄ ⁄‰ «·«·€«¡ »‰Ã«Õ"
    myLoadHeader
    Handlecontrols
End If
End Sub
Private Function cancelDelete() As Boolean
Dim con As New ADODB.Connection
openCon con
On Error GoTo myerror:
con.Execute "UPDATE FILE6_90H " & _
            " SET FILE6_90H.DELORDER_DATE = NULL," & _
            " FILE6_90H.NOTES = NULL" & _
            " FROM FILE6_90H WHERE DOC_NO = " & MyParn(sDoc_no)
cancelDelete = True
Finally:
closeCon con
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GoTo Finally
End Function
Private Sub cmdExit_Click()
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
Private Sub cmdReturn_Click()
Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "select TOP 1 FILE6_20H.BRANCH," & _
          " FILE6_20H.PRINTED," & _
          " FILE6_20H.ISINVOICE," & _
          " FILE6_20H.TOTAL_ITEM," & _
          " FILE6_20H.DATE, " & _
          " FILE6_20H.ISRET" & _
          " FROM FILE6_20H " & _
          " WHERE FILE6_20H.DOC_NO = " & MyParn(xSales_doc.Caption)

Set loctable = myRs(cString)
If loctable.EOF Then
    MsgBox "›« Ê—… €Ì— „”Ã·…"
    Exit Sub
End If

If loctable!branch & "" = "00" Then
    MsgBox "›« Ê—… ›—⁄ —∆Ì”Ì"
    Exit Sub
End If

If Not loctable!PRINTED Then
    MsgBox "›« Ê—… „› ÊÕ…"
    Exit Sub
End If

If loctable!ISINVOICE Then
    MsgBox "›« Ê—… „Ã„⁄…"
    Exit Sub
End If

If loctable!TOTAL_ITEM <= 0 Then
    If loctable!TOTAL_ITEM = 0 Then
        MsgBox "›« Ê—… »œÊ‰ ﬁÌ„…"
    End If
    If loctable!TOTAL_ITEM < 0 Then
        MsgBox "›« Ê—… „— Ã⁄"
    End If
    Exit Sub
End If

Set oSalesRefund.myForm = Me
oSalesRefund.sDoc_no = xSales_doc.Caption
oSalesRefund.Show 1
End Sub
Private Sub cmdSave_Click()
If MsgBox("Õ›Ÿ", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
If myreplace Then
    myload
    Inform " „ «·Õ›Ÿ »‰Ã«Õ"
End If
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "[NAME]", addstring(xName.text))
aInsert = AddFlag(aInsert, "[PHONE]", addstring(xphone.text))
aInsert = AddFlag(aInsert, "[CITY]", addstring(xcity.text))
aInsert = AddFlag(aInsert, "[street]", addstring(xStreet.text))
aInsert = AddFlag(aInsert, "[SHIP_DATE]", addstring(xship_date.text))
aInsert = AddFlag(aInsert, "[NOTE_MAIN]", addstring(xNote_main.text))
aInsert = AddFlag(aInsert, "[NOTE_BRANCH]", addstring(xNote_branch.text))
aInsert = AddFlag(aInsert, "[DELORDER_DATE]", addDate(xDelOrder_Date.text))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.text))
On Error GoTo myerror
openCon con
con.Execute addUpdate(aInsert, "FILE6_90H", "DOC_NO = " & MyParn(sDoc_no))
myreplace = True
Finally:
closeCon con
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GoTo Finally
End Function
Private Sub Handlecontrols()
xNote_main.Enabled = cBranch = "00"
xNote_branch.Enabled = cBranch <> "00"
cmdSave.Enabled = xClosed.Value = 0
fmAddInvoice.Visible = xSales_doc.Caption <> "" And Trim(xDelOrder_Date.text) = "" And cBranch = "00"
cmdSend.Enabled = Trim(xDelOrder_Date.text) = "" And xSales_doc.Caption = "" And cBranch <> "00"
xDelOrder_Date.Enabled = xSales_doc.Caption = ""
cmdDelCancel.Visible = IsDate(xDelOrder_Date.text)
xNotes.Enabled = xSales_doc.Caption = ""
End Sub

Private Sub cmdsend_Click()
Dim oOnlineCheck As New OnlineCheck
Set oOnlineCheck.myForm = Me
oOnlineCheck.sDoc_no = sDoc_no
oOnlineCheck.Show 1
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
If sStore = "" Then
    sStore = rsValue("SELECT CODE FROM FILE0_40 WHERE ONLINE = 1 ORDER BY CODE")
End If
myload
myLoadCount
End Sub
Private Sub myLoadGrd()
Dim cString As String
On Error GoTo myerror
With grid1
    cString = " SELECT FILE6_90.ITEM," & _
              " FILE1_10.MODELFACT0," & _
              " FACT.DESCA," & _
              " FILE1_10.desca," & _
              " FILE1_10.COLOR," & _
              " FILE1_10.SCAL," & _
              " FILE6_90.QUANT," & _
              " FILE6_90.PRICE," & _
              " FILE6_90.TOTAL " & _
              " FROM  FILE6_90 " & _
              " INNER JOIN FILE6_90H ON FILE6_90.DOC_NO = FILE6_90H.DOC_NO " & _
              " INNER  JOIN FILE1_10 ON FILE6_90.ITEM = FILE1_10.ITEM" & _
              " INNER  JOIN  FACT ON FILE1_10.code = FACT.CODE" & _
              " WHERE FILE6_90H.DOC_NO =  " & MyParn(sDoc_no)
    Set grid1.DataSource = myRs(cString)
End With

fixGrd

grid1.Cell(flexcpAlignment, 0, 0, grid1.Rows - 1, grid1.Cols - 1) = 7
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Public Sub myload()
myLoadHeader
myLoadCount
myLoadGrd
Handlecontrols
End Sub
Private Sub myLoadHeader()
Dim loctable As New ADODB.Recordset
Set loctable = myRs("SELECT FILE6_90H.*,CASE WHEN V.ORDER_NO IS NULL THEN 1 ELSE 0 END AS CLOSED  " & _
                    " FROM FILE6_90H " & _
                    " LEFT JOIN vw_online_orders_open AS v ON FILE6_90H.DOC_NO = v.ORDER_NO" & _
                    " WHERE DOC_NO = " & MyParn(sDoc_no))
If Not loctable.EOF Then
    xName.text = loctable!Name & ""
    xphone.text = loctable!phone & ""
    xcity.text = loctable!city & ""
    xStreet.text = loctable!STREET & ""
    xship_date.text = myFormat_p(loctable!ship_date)
    xNotes.text = loctable!NOTES & ""
    xNote_branch.text = loctable!note_branch & ""
    xNote_main.text = loctable!note_main & ""
    xDelOrder_Date.text = myFormat_p(loctable!delorder_date)
    xSales_doc.Caption = loctable!sales_Doc & ""
    xClosed.Value = loctable!CLOSED
End If
loctable.Close
Set loctable = Nothing
End Sub
Sub fixGrd()
Dim bShowColAll As Boolean, bShowCol As Boolean
With grid1
        
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "„ÊœÌ· «·„’‰⁄"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "«·’‰›"
    .TextMatrix(0, 4) = "«··Ê‰"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«·ﬂ„Ì…"
    .TextMatrix(0, 7) = "«·”⁄—"
    .TextMatrix(0, 8) = "«·≈Ã„«·Ì"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1500
    .ColWidth(2) = 3000
    .ColWidth(3) = 5000
    .ColWidth(4) = 3000
    .ColWidth(5) = 1200
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1400
                           
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
Public Function myReplaceRefund(pDiscount_add As Double, pDiscount_total As Double) As Boolean

Dim con As New ADODB.Connection
openCon con


Dim aInsert As Variant
With oSalesRefund
Dim sdoc_no_new As String
sdoc_no_new = IncData("FILE6_90BH", "DOC_NO", con)
If sdoc_no_new = "" Then sdoc_no_new = "1"

aInsert = AddFlag(Empty, "DOC_NO", sdoc_no_new)
aInsert = AddFlag(aInsert, "[DISCOUNT]", -1 * Val(.xdiscount_ret.Caption))
aInsert = AddFlag(aInsert, "[ORDER_NO]", addstring(sDoc_no))
aInsert = AddFlag(aInsert, "[DATE]", addDate(Date))
aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cUserName))
aInsert = AddFlag(aInsert, "[USER_IP]", addstring(GetComputerName))
End With

con.BeginTrans
On Error GoTo myerror
con.Execute addInsert(aInsert, "FILE6_90BH")

With oSalesRefund.grid1
For i = 1 To .Rows - 2
    aInsert = AddFlag(Empty, "ITEM", addstring(.TextMatrix(i, 1)))
    aInsert = AddFlag(aInsert, "QUANT", -1 * .ValueMatrix(i, 8))
    aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 9))
    aInsert = AddFlag(aInsert, "PRICE_C", .ValueMatrix(i, 7))
    aInsert = AddFlag(aInsert, "COST", .ValueMatrix(i, 15))
    aInsert = AddFlag(aInsert, "MAN", addstring(.TextMatrix(i, 15 + 1)))
    aInsert = AddFlag(aInsert, "USER_IP", addstring(GetComputerName))
    aInsert = AddFlag(aInsert, "OFFER_NO", addstring(.TextMatrix(i, 11)))
    aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(.TextMatrix(i, 12)))
    aInsert = AddFlag(aInsert, "DISCOUNT_OFFER", -1 * .ValueMatrix(i, 13))
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(sDoc_no))
    con.Execute addInsert(aInsert, "FILE6_90B")
Next
End With
con.CommitTrans
myReplaceRefund = True

MsgBox " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"

Unload oSalesRefund

Finally:
closeCon con
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub Form_Resize()
Dim nHeight As Long
nHeight = Me.Height - (grid1.Top + 1500)
If nHeight > 3000 Then
    grid1.Height = nHeight
End If
End Sub
Public Sub myLoadCount()
Dim nCount As Long
nCount = Val(rsValue("SELECT COUNT(*) FROM FILE6_90BH WHERE ORDER_NO = " & MyParn(sDoc_no)) & "")
If nCount > 0 Then
    cmdAddInvoice.Caption = ArbString("«÷«›… ÿ·»Ì…" & " (" & nCount & ")")
Else
    cmdAddInvoice.Caption = "«÷«›… ÿ·»Ì…"
End If
End Sub

Private Sub xNote_branch_GotFocus()
myGotFocus xNote_branch
End Sub
Private Sub xNote_branch_LostFocus()
myLostFocus xNote_branch
End Sub
Private Sub xNote_main_GotFocus()
myGotFocus xNote_main
End Sub
Private Sub xNote_main_LostFocus()
myLostFocus xNote_main
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xDelOrder_Date_GotFocus()
myGotFocus xDelOrder_Date
End Sub
Private Sub xDelOrder_Date_LostFocus()
myLostFocus xDelOrder_Date
myValidDate xDelOrder_Date
End Sub
Private Sub xPhone_GotFocus()
myGotFocus xphone
End Sub
Private Sub XPHONE_LostFocus()
myLostFocus xphone
End Sub
Private Sub xcity_GotFocus()
myGotFocus xcity
End Sub
Private Sub xcity_LostFocus()
myLostFocus xcity
End Sub
Private Sub xship_date_GotFocus()
myGotFocus xship_date
End Sub
Private Sub xship_date_LostFocus()
myLostFocus xship_date
myValidDate xship_date
End Sub
Private Sub xStreet_GotFocus()
myGotFocus xStreet
End Sub
Private Sub xStreet_LostFocus()
myLostFocus xStreet
End Sub
Private Sub xName_GotFocus()
myGotFocus xName
End Sub
Private Sub xname_LostFocus()
myLostFocus xName
End Sub

