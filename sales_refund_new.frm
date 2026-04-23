VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form sales_refundfrm 
   BackColor       =   &H80000005&
   Caption         =   "„— Ã⁄«  «’‰«›"
   ClientHeight    =   10200
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   20370
   FillColor       =   &H003F3E12&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   -45
      Width           =   4560
      Begin VB.TextBox xItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   3615
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·’‰›"
         Height          =   285
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   630
         Width           =   645
      End
      Begin VB.Label xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   585
         Width           =   3615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·»«—ﬂÊœ"
         Height          =   285
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   4635
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   -45
      Width           =   10905
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰”»… «·⁄—Ê÷"
         Height          =   285
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label xRate_net 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   585
         Width           =   1365
      End
      Begin VB.Label xDiscount_offer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Œ’„ «·⁄—Ê÷"
         Height          =   330
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·›« Ê—…"
         Height          =   330
         Left            =   10125
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   270
         Width           =   645
      End
      Begin VB.Label xdoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8415
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·›« Ê—…"
         Height          =   330
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·›—⁄"
         Height          =   330
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   270
         Width           =   555
      End
      Begin VB.Label xBranch 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   1680
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·≈Ã„«·Ì"
         Height          =   330
         Left            =   10125
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   585
         Width           =   690
      End
      Begin VB.Label xTotal_inv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8415
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   585
         Width           =   1590
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Œ’„ «·›« Ê—…"
         Height          =   330
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   630
         Width           =   1185
      End
      Begin VB.Label xDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   1500
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰”»… Œ’„ «·›« Ê—…"
         Height          =   285
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   870
      End
      Begin VB.Label xDiscount_Rate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   1680
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3390
      Left            =   45
      TabIndex        =   0
      Top             =   1035
      Width           =   15495
      _cx             =   27331
      _cy             =   5980
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
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   -765
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   -765
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   4065
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4455
      Width           =   15540
      _cx             =   27411
      _cy             =   7170
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
      GridColorFixed  =   6841886
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
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   -45
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   8460
      Width           =   6045
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "sales_refund_new.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSend 
         Height          =   555
         Left            =   4320
         TabIndex        =   5
         Top             =   180
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   979
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
         Picture         =   "sales_refund_new.frx":2323
         Caption         =   " ”ÃÌ· „— Ã⁄"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_refund_new.frx":4977
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   555
         Left            =   2655
         TabIndex        =   19
         Top             =   180
         Width           =   1590
         _ExtentX        =   2805
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
         Picture         =   "sales_refund_new.frx":6E17
         Caption         =   "«Œ Ì«— «·ﬂ·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_refund_new.frx":92ED
      End
      Begin Threed.SSCommand cmdDelAll 
         Height          =   555
         Left            =   1215
         TabIndex        =   35
         Top             =   180
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   979
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
         Picture         =   "sales_refund_new.frx":B5E3
         Caption         =   "Õ–› «·ﬂ·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_refund_new.frx":DC7D
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   8460
      Width           =   9600
      Begin VB.Label xQuant_Ret 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7155
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬂ„Ì… „— Ã⁄"
         Height          =   285
         Left            =   8505
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì"
         Height          =   285
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label xTotal_Ret 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Œ’„"
         Height          =   285
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   315
         Width           =   510
      End
      Begin VB.Label xDiscount_ret 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "’«›Ì"
         Height          =   285
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label xTotal_ret_Net 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   270
         Width           =   1365
      End
   End
   Begin VB.Label xPhone 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "sales_refundfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
'Public sDoc_no_ret As String
Public sOnline_doc As String
Public sid_lading
Public myForm As Form
Dim bStopCell As Boolean
Dim con As New ADODB.Connection
Dim conServer As New ADODB.Connection
Private Sub CMD_SEND_Click()
'If Not myValid Then Exit Sub
'Dim sNewDoc As String
'sNewDoc = myReplacereturn
'If sNewDoc = "" Then Exit Sub
'MsgBox " „ «· —ÕÌ· «·Ì „” ‰œ „»Ì⁄«  —ﬁ„ " & sNewDoc
'myform.myproc_ret sNewDoc
End Sub

Private Sub CmdDelAll_Click()
grid1.Rows = 1
CalcTotals
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSelectAll_Click()
'Dim i As Long
'grid1.Rows = 1
'myloadgrd2
'For i = 1 To grid2.Rows - 2
'    addRow i
'Next
'
'Fixgrd2
'fixGrd
'
'CalcTotals
grid1.Rows = 1
CalcTotals
'myLoadGrd
addRow
fixGrdColor
End Sub

Private Sub cmdsend_Click()
myForm.myReplaceRefund
End Sub

Private Sub Form_Load()
openCon con

Set grid1.DataSource = data1
Set grid2.DataSource = data2

myLoadInv

myloadgrd2

fixGrd
CalcTotals

CellPos 13, 0, grid1.Cols - 1
End Sub
Private Sub myLoadInv()
Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "SELECT FILE6_20H.DATE," & _
          "FILE6_20H.TOTAL_ITEM," & _
          "FILE6_20H.DISCOUNT," & _
          "FILE6_20H.DISCOUNT_RATE," & _
          "FILE6_20H.RATE_NET," & _
          "FILE6_20H.DISCOUNT_OFFER," & _
          "FILE0_40.DESCA AS BRANCH," & _
          "FILE6_20H.PHONE" & _
          " FROM FILE6_20H " & _
          " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE" & _
          " WHERE FILE6_20H.DOC_NO = " & MyParn(sDoc_no) & _
          "UNION ALL " & _
          "SELECT FR6_20H.DATE," & _
          "FR6_20H.TOTAL_ITEM," & _
          "FR6_20H.DISCOUNT," & _
          "FR6_20H.DISCOUNT_RATE," & _
          "FR6_20H.RATE_NET," & _
          "FR6_20H.DISCOUNT_OFFER," & _
          "BRANCH_FR.DESCA AS BRANCH," & _
          "FR6_20H.PHONE" & _
          " FROM FR6_20H " & _
          " INNER JOIN BRANCH_FR ON FR6_20H.STORE = BRANCH_FR.CODE" & _
          " WHERE FR6_20H.DOC_NO = " & MyParn(sDoc_no)

          
Set loctable = cmd(cString, con).Execute
If Not loctable.EOF Then
    xDoc_no.Caption = sDoc_no
    xTotal_inv.Caption = Myvalue(loctable!TOTAL_ITEM)
    XBRANCH.Caption = loctable!branch & ""
    xDiscount.Caption = Myvalue(loctable!discount & "")
    xDiscount_offer.Caption = Myvalue(loctable!discount_offer & "")
    xDate.Caption = myFormat_p(loctable!Date)
    xDiscount_Rate.Caption = Myvalue(Round(loctable!discount_Rate * 100, 2))
    xRate_net.Caption = Format(Round(loctable!Rate_net, 4), "Percent")
    If IsValidMobile(loctable!phone & "") Then
        xPhone.Caption = loctable!phone
    End If
    xRate_net.Tag = Val(loctable!Rate_net & "")
End If
End Sub

Private Sub myloadgrd2()
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DOC_NO", sDoc_no)
Set data2.Recordset = mycmd("dbo.sp_SALES_REFUND", con, adStoredProc, aPrm)
Fixgrd2
End Sub
Sub Fixgrd2()
With grid2
    .RowHeight(0) = 600

    .TextMatrix(0, 0) = "„"
    .TextMatrix(0, 1) = "ﬂÊœ"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "„ÊœÌ· «·„’‰⁄"
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«··Ê‰"
    .TextMatrix(0, 7) = "”⁄— «·„” Â·ﬂ"
    .TextMatrix(0, 8) = "«·ﬂ„Ì…"
    .TextMatrix(0, 9) = "«·”⁄—"
    .TextMatrix(0, 10) = "«·≈Ã„«·Ì"
    
    .TextMatrix(0, 11) = "„— Ã⁄"
    .TextMatrix(0, 12) = "«·»«ﬁÌ"
    .TextMatrix(0, 13) = "«· ﬂ·›…"
    
    .TextMatrix(0, 14) = "—ﬁ„ «·⁄—÷"
    .TextMatrix(0, 15) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 16) = "Œ’„ «·„” ‰œ"
    .TextMatrix(0, 17) = "»Ì«‰ «·⁄—÷"
    
    .TextMatrix(0, 18) = "≈”„ «·»«∆⁄"
    .TextMatrix(0, 19) = "»Ì«‰ «·⁄—÷"
    .TextMatrix(0, 20) = "Flag"
    .TextMatrix(0, 21) = "„— Ã⁄"
    
    .ColWidth(0) = 600
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1200
    .ColWidth(4) = 2500
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 800
    .ColWidth(8) = 700
    .ColWidth(9) = 900
    .ColWidth(10) = 1000
    .ColWidth(11) = 1000
    .ColWidth(12) = 700
    .ColWidth(13) = 1000
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    .ColWidth(16) = 1000
    .ColWidth(17) = 1000
    .ColWidth(18) = 1000
    .ColWidth(19) = 1600

'    .ColHidden(10 +1) = True
'    .ColHidden(12 +1) = True
'    .ColHidden(13 +1) = True
'    .ColHidden(14 +1) = True
'    .ColHidden(15 +1) = True
'    .ColHidden(16 +1) = True
'    .ColHidden(17 +1) = True
'    .ColHidden(19 +1) = True
    
    .ColHidden(11) = True
    .ColHidden(13) = True
    .ColHidden(14) = True
    .ColHidden(15) = True
    .ColHidden(16) = True
    .ColHidden(17) = True
    
    .ColHidden(.Cols - 2) = True
    .ColHidden(.Cols - 1) = True

    .ColWidth(grid2.Cols - 1) = 1000
    For i = 1 To .Rows - 1
        If .ValueMatrix(i, 7) <> .ValueMatrix(i, 9) Then
            .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
            .Cell(flexcpFontUnderline, i, 1, i, .Cols - 1) = vbRed
        End If
       
       If .ValueMatrix(i, 14) > 0 Then
            .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = &H515017
       End If
    Next
    
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    
    For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = i
    Next
    .MergeCells = flexMergeFree
    .MergeCol(19) = True
End With
End Sub
Private Sub CalcTotals()
grid1.Subtotal flexSTClear
grid2.Subtotal flexSTClear

cmdDelAll.Enabled = grid1.Rows > 1
cmdSend.Enabled = grid1.Rows > 1

With grid2
Dim i As Long
Dim sGroup As String
Dim bGroup As Boolean
For i = 1 To .Rows - 1
    nFound = grid1.FindRow(.TextMatrix(i, .Cols - 2), , grid1.Cols - 1)
    If nFound <> -1 Then
         .TextMatrix(i, 11) = .ValueMatrix(i, .Cols - 1) + grid1.ValueMatrix(nFound, 8)
         .TextMatrix(i, 12) = .ValueMatrix(i, 8) - .ValueMatrix(i, 11)
    Else
         .TextMatrix(i, 11) = .ValueMatrix(i, .Cols - 1)
         .TextMatrix(i, 12) = .ValueMatrix(i, 8) - .ValueMatrix(i, 11)
    End If
    
    If .ValueMatrix(i, 12) = 0 Then
        .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &HC0C0FF
    ElseIf .ValueMatrix(i, 12) > 0 And .ValueMatrix(i, 12) < .ValueMatrix(i, 8) Then
        .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &H80FFFF
    ElseIf .ValueMatrix(i, 11) = .ValueMatrix(i, .Cols - 1) Then
        If .ValueMatrix(i, 14) > 0 Then
            .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = &H515017
            If sGroup <> .TextMatrix(i, 14) Then
                sGroup = .TextMatrix(i, 14)
                bGroup = Not bGroup
            End If
            .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = RGB(IIf(bGroup, 255, 200), 255, 210)
        Else
            .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbWhite
        End If
    End If
Next
.SubtotalPosition = flexSTBelow
.Subtotal flexSTSum, -1, 8, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
.Subtotal flexSTSum, -1, 10, "#.##", &HE0E0E0, , True, "«·≈Ã„«·Ì"
.Subtotal flexSTSum, -1, 12, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
End With

With grid1
    .Subtotal flexSTClear
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 8, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, 10, "#.##", &HE0E0E0, , True, "«·≈Ã„«·Ì"
End With

With grid1
If grid1.Rows > 1 Then
    Dim nDiscount As Long
    Dim nDiscountOffer As Long
    Dim nOffer As Long
    
    
    If grid1.ValueMatrix(grid1.Rows - 1, 8) = grid2.ValueMatrix(grid2.Rows - 1, 8) And grid1.ValueMatrix(grid1.Rows - 1, 10) = grid2.ValueMatrix(grid2.Rows - 1, 10) Then
        xDiscount_ret.Caption = xDiscount.Caption
    Else
        For i = 1 To .Rows - 2
            If .ValueMatrix(i, 11) <> 0 Then
                If nOffer <> .ValueMatrix(i, 11) Then
                    nOffer = .ValueMatrix(i, 11)
                    nDiscountOffer = nDiscountOffer + .ValueMatrix(i, 13)
                End If
            End If
        Next
        nDiscount = Round(.ValueMatrix(.Rows - 1, 10) * Val(xRate_net.Tag), 2)
        xDiscount_ret.Caption = nDiscountOffer + nDiscount
    End If
    xQuant_Ret.Caption = grid1.ValueMatrix(grid1.Rows - 1, 8)
    xTotal_Ret.Caption = grid1.ValueMatrix(grid1.Rows - 1, 10)
    xTotal_ret_Net.Caption = grid1.ValueMatrix(grid1.Rows - 1, 10) - Val(xDiscount.Caption)
Else
    xQuant_Ret.Caption = ""
    xTotal_Ret.Caption = ""
    xTotal_ret_Net.Caption = ""
    xDiscount_ret.Caption = ""
End If
End With
End Sub
Sub fixGrd()
With grid1
    .RowHeight(0) = 600
    
    .TextMatrix(0, 0) = "„"
    .TextMatrix(0, 1) = "ﬂÊœ"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "„ÊœÌ· «·„’‰⁄"
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«··Ê‰"
    .TextMatrix(0, 7) = "”⁄— «·„” Â·ﬂ"
    .TextMatrix(0, 8) = "«·ﬂ„Ì…"
    .TextMatrix(0, 9) = "«·”⁄—"
    .TextMatrix(0, 10) = "«·≈Ã„«·Ì"
                
    .TextMatrix(0, 11) = "—ﬁ„ «·⁄—÷"
    .TextMatrix(0, 12) = "„” ‰œ «·⁄—÷"
    .TextMatrix(0, 13) = "Œ’„ «·⁄—÷"
    .TextMatrix(0, 14) = "»Ì«‰ «·⁄—÷"
    .TextMatrix(0, 15) = "«· ﬂ·›…"
    .TextMatrix(0, 16) = "Flag"
        
    .ColWidth(0) = 600
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1500
    .ColWidth(4) = 3000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1200
    .ColWidth(9) = 750
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    .ColWidth(12) = 1200
    .ColWidth(13) = 1200
    .ColWidth(14) = 1600
    
    .ColHidden(11) = True
    .ColHidden(12) = True
    .ColHidden(13) = True
    .ColHidden(15) = True
    .ColHidden(16) = True
    
    Dim i As Long
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    
    .MergeCells = flexMergeFree
    .MergeCol(14) = True
    '.MergeCells = flexMergeRestrictRows
    '.MergeCells = flexMergeRestrictColumns
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set sales_refundfrm = Nothing
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
grid1.TextMatrix(Row, 10) = mRound(grid1.ValueMatrix(Row, 8) * grid1.ValueMatrix(Row, 9))
CalcTotals
End Sub

Private Sub grid1_EnterCell()
If grid1.col = 8 And grid1.ValueMatrix(grid1.Row, 11) = 0 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If Val(grid1.EditText) <= 0 Then
    Cancel = True
    Exit Sub
End If

Dim nFound As Long
nFound = grid2.FindRow(grid1.TextMatrix(Row, grid1.Cols - 1), , grid2.Cols - 2)
If grid2.ValueMatrix(nFound, 8) - grid2.ValueMatrix(nFound, grid2.Cols - 1) < Val(grid1.EditText) Then
    Cancel = True
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf col < 9 Then
    grid1.col = 9
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 9, 9)
    grid1.ShowCell grid1.Row, 9
Else
    grid1.Select Row, col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 4 Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'If grid1.Col = 4 Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    sMsg = IIf(grid1.ValueMatrix(grid1.Row, 11) = 0, "Õ–› „‰ «·’‰› ?, Â· «‰  „Ê«›ﬁ ø", "Õ–› «·⁄—÷ " & grid1.TextMatrix(grid1.Row, 14) & " »«·ﬂ«„·")
    If MsgBox(sMsg, vbOKCancel) = vbOK Then
        If grid1.ValueMatrix(grid1.Row, 11) = 0 Then
            myRemove grid1.Row
        Else
            nOffer_no = grid1.ValueMatrix(grid1.Row, 11)
            grid1.Redraw = flexRDNone
            For i = grid1.Rows - 2 To 1 Step -1
                If grid1.ValueMatrix(i, 11) = nOffer_no Then
                    grid1.RemoveItem i
                End If
            Next
            fixGrdColor
            grid1.Redraw = flexRDBuffered
        End If
        CalcTotals
        bStopCell = True
    End If
End If
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
End Sub
Private Function MYVALID() As Boolean
Dim i As Long
Dim nFound As Long

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«› „”Ã·…"
    Exit Function
End If

With grid2
For i = 1 To .Rows - 2
    nFound = grid1.FindRow(.TextMatrix(i, grid2.Cols - 2), , grid1.Cols - 1)
    If nFound = -1 Then
        If .ValueMatrix(i, 14) > 0 Then
            If MsgBox("«·’‰› " & .TextMatrix(i, 1) & " ·„ Ì”Ã· ›Ï «·›« Ê—…", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
        End If
    ElseIf .TextMatrix(i, 14) < 0 Then
        MsgBox "«·’‰› " & .TextMatrix(i, 1) & " «ﬂ»— „‰ «·›« Ê—… » " & Abs(.ValueMatrix(i, 14))
        Exit Function
    ElseIf .TextMatrix(i, 14) > 0 Then
        If MsgBox("«·’‰› " & .TextMatrix(i, 1) & "  „ »ﬁÌ »ﬂ„Ì… " & .ValueMatrix(i, 14), vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
    End If
Next
MYVALID = True
End With
End Function
Private Sub Grid2_dblClick()
AddInvItem grid2.Row
End Sub
Private Sub AddInvItem(Row As Long, Optional nquant As Long = -1)
With grid2
Dim i As Long
If Row = 0 Or Row = grid2.Rows - 1 Then Exit Sub
If grid2.ValueMatrix(Row, 14) = 0 Then
    addRow Row, nquant
Else
    If grid2.TextMatrix(Row, 12) = 0 Then Exit Sub
    If MsgBox("«÷«›… ⁄—÷ " & grid2.TextMatrix(Row, 19), vbOKCancel + vbDefaultButton1) <> vbOK Then
        Exit Sub
    End If
    
    Dim nOffer As String
    nOffer = grid2.TextMatrix(Row, 14)
    nFound = grid2.FindRow(nOffer, , 14)
    For i = nFound To grid2.Rows - 2
        If grid2.ValueMatrix(i, 14) <> nOffer Then Exit For
        addRow i
    Next
End If
fixGrdColor
grid1.ShowCell grid1.Rows - 1, 1
End With
End Sub
Private Sub addRow(Optional pRow As Long = -1, Optional nquant As Long = -1)
Dim i As Long
Dim Row As Long
With grid1
For Row = IIf(pRow = -1, 1, pRow) To IIf(pRow = -1, grid2.Rows - 2, pRow)
    If grid2.ValueMatrix(Row, 12) > 0 Then
        nFound = grid1.FindRow(grid2.TextMatrix(Row, grid2.Cols - 2), , .Cols - 1)
        If nFound <> -1 Then
            grid1.Redraw = flexRDNone
            grid1.RemoveItem nFound
            grid1.Redraw = flexRDBuffered
        End If
        grid1.AddItem ""
        For i = 1 To 10
            If i = 8 And nquant > -1 Then
                grid1.TextMatrix(grid1.Rows - 1, i) = nquant
            Else
                grid1.TextMatrix(grid1.Rows - 1, i) = grid2.TextMatrix(Row, i)
            End If
        Next
        grid1.TextMatrix(grid1.Rows - 1, 11) = grid2.TextMatrix(Row, 14)
        grid1.TextMatrix(grid1.Rows - 1, 12) = grid2.TextMatrix(Row, 15)
        grid1.TextMatrix(grid1.Rows - 1, 13) = grid2.TextMatrix(Row, 16)
        grid1.TextMatrix(grid1.Rows - 1, 15) = grid2.TextMatrix(Row, 13)
        grid1.TextMatrix(grid1.Rows - 1, 14) = grid2.TextMatrix(Row, 19)
        grid1.TextMatrix(grid1.Rows - 1, grid1.Cols - 1) = grid2.TextMatrix(Row, grid2.Cols - 2)
    End If
Next
CalcTotals
fixSerial
End With
End Sub
Private Sub xItem_Change()
xdesca.Caption = ""
End Sub

Private Sub xitem_GotFocus()
myGotFocus xItem
End Sub
Private Sub xITEM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(xItem.text) = "" Then Exit Sub
    myGotFocus xItem
    
    Dim cm As New ADODB.command
    Set cm = cmd("dbo.sp_item", con, adStoredProc, AddFlag(Empty, "BARCODE_FIND", xItem.text))
    cm.Execute

    If IsNull(cm.Parameters("@ITEM").value) Then
        sMsg = "ﬂÊœ €Ì— ’«·Õ"
        Exit Sub
    End If

    If cm.Parameters("@ITEM").value <> xItem.text Then
        xItem.text = cm.Parameters("@ITEM").value
    End If
    
    xdesca.Caption = cm.Parameters("@DESCA").value
    Dim i As Long
    For i = 1 To grid2.Rows - 2
        If grid2.TextMatrix(i, 1) = Trim(xItem.text) And grid2.ValueMatrix(i, 12) > 0 Then
            AddInvItem i
            Exit Sub
        End If
    Next
End If
End Sub

Private Sub xitem_LostFocus()
myLostFocus xItem
End Sub
Private Sub fixSerial()
Dim i As Long
For i = 1 To grid1.Rows - 2
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub fixGrdColor()
With grid1
Dim bGroup As Boolean
Dim sGroup As String
For i = 1 To .Rows - 1
    If .ValueMatrix(i, 7) <> .ValueMatrix(i, 9) Then
        .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
        .Cell(flexcpFontUnderline, i, 1, i, .Cols - 1) = vbRed
    End If
   
    If .ValueMatrix(i, 11) > 0 Then
        .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = &H515017
        If sGroup <> .TextMatrix(i, 11) Then
            sGroup = .TextMatrix(i, 11)
            bGroup = Not bGroup
        End If
        .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = RGB(IIf(bGroup, 255, 200), 255, 210)
    Else
        .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbWhite
    End If
Next
End With
End Sub
