VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form online_Stage_invoices 
   BackColor       =   &H00FFFFFF&
   Caption         =   " ÃÂÌ“ ÿ·»Ì…"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   14355
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   735
      ScaleWidth      =   14355
      TabIndex        =   6
      Top             =   8385
      Width           =   14355
      Begin Threed.SSCommand cmdStage 
         Height          =   555
         Index           =   5
         Left            =   9765
         TabIndex        =   17
         Top             =   90
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
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
         Caption         =   "«—”«· ··«·€«¡"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStage 
         Height          =   555
         Index           =   2
         Left            =   12015
         TabIndex        =   11
         Top             =   90
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
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
         Caption         =   "»œ« «· ÃÂÌ“"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   3
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   0
         Width           =   3705
         Begin Threed.SSCommand cmdExit 
            Height          =   555
            Left            =   45
            TabIndex        =   8
            TabStop         =   0   'False
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
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "online_stage_invoices.frx":0000
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdPrint 
            Height          =   555
            Left            =   1215
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   135
            Width           =   1185
            _ExtentX        =   2090
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
            Picture         =   "online_stage_invoices.frx":2323
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "online_stage_invoices.frx":4699
         End
         Begin Threed.SSCommand cmdExcel 
            Height          =   555
            Left            =   2430
            TabIndex        =   10
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
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
            Picture         =   "online_stage_invoices.frx":681C
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
      End
      Begin Threed.SSCommand cmdStage 
         Height          =   555
         Index           =   3
         Left            =   12015
         TabIndex        =   12
         Top             =   90
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
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
         Caption         =   " „ «· ÃÂÌ“"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStage 
         Height          =   555
         Index           =   4
         Left            =   9765
         TabIndex        =   13
         Top             =   90
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
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
         Caption         =   "„‘ﬂ·… ›Ï «·ÿ·»Ì…"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStage 
         Height          =   555
         Index           =   6
         Left            =   7515
         TabIndex        =   18
         Top             =   90
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
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
         Caption         =   "«—”«· ·· ⁄œÌ·"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdInvoice 
         Height          =   555
         Left            =   12015
         TabIndex        =   19
         Top             =   90
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
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
         Caption         =   "⁄„· ›« Ê—…"
         TagVariant      =   "√Œ «— «·„‘—Ê⁄"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   14145
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·Œ’„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label xDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   7065
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   540
         Width           =   2130
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Õ«·… «·ÿ·»Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label xStage 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   7065
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   2130
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«· «—ÌŒ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13140
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   825
      End
      Begin VB.Label xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   10935
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   2130
      End
      Begin VB.Label xDoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   10935
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   2130
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·ÿ·»Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13140
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -1485
      Top             =   5490
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4965
      Left            =   135
      TabIndex        =   0
      Top             =   1125
      Width           =   14145
      _cx             =   24950
      _cy             =   8758
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grdError 
      Height          =   2220
      Left            =   225
      TabIndex        =   14
      Top             =   6120
      Width           =   14145
      _cx             =   24950
      _cy             =   3916
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "online_Stage_invoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
Public sManCode As String
Public bInvoice As Boolean
Public sMsg As String
Public myForm As Form
Dim aHeader(2) As String
Dim bError As Boolean
Dim aRow As Variant, aXl As Variant
Private Sub cmdExel_Click()
ToFileExelNew grid1, , , aXl, , 1
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdPrint_Click()
'Dim aRow(0) As Variant
'aRow(0) = AddFlag(Empty, "row", 1)
'aRow(0) = AddFlag(aRow(0), "col", 0)
'aRow(0) = AddFlag(aRow(0), "cols", IIf(retFlag(aData, "ITEM") = "", 5, 4))
'printGrdNew.sDeviceType = "1"
printGrdNew.bIgPreview = True
printGrdNew.doprint grid1, 0.85, -3, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 4), retHeader(aHeader, 6, 4), False, False, 10, , aRow
Set printGrdNew = Nothing
'printGrdNew.Show 1
End Sub
Private Sub cmdStage_Click(Index As Integer)
If updateStage(Index) Then
    myForm.myload
    Unload Me
End If
End Sub
Private Sub Form_Load()
Dim cString As String
cString = "SELECT FILE6_90H.MAN," & _
          "FILE6_90H.DISCOUNT," & _
          " FILE6_90H.STAGE," & _
          " STAGES_CODES.DESCA AS STAGE_DESCA," & _
          " FILE6_25.DESCA AS MAKE_DESCA" & _
          " FROM FILE6_90H " & _
          " INNER JOIN STAGES_CODES ON FILE6_90H.STAGE = STAGES_CODES.CODE" & _
          " LEFT JOIN  FILE6_25 ON FILE6_90H.MAN = FILE6_25.CODE" & _
          " WHERE FILE6_90H.DOC_NO = " & MyParn(sDoc_no)

On Error GoTo myerror
Dim loctable As New ADODB.Recordset
Set loctable = myRs(cString)

xStage.Caption = loctable!stage_Desca
xStage.Tag = loctable!Stage
xDiscount.Caption = loctable!discount

myload
myloadGrdError
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrd()
With grid1
        
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "„ÊœÌ·"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "«·’‰›"
    .TextMatrix(0, 4) = "«··Ê‰"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«·ﬂ„Ì…"
    .TextMatrix(0, 7) = "«·”⁄—"
    .TextMatrix(0, 8) = "«·≈Ã„«·Ì"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1500
    .ColWidth(2) = 2000
    .ColWidth(3) = 4000
    .ColWidth(4) = 2500
    .ColWidth(5) = 800
    .ColWidth(6) = 800
    .ColWidth(7) = 1000
    .ColWidth(8) = 1200
    
    '.ColHidden(1) = True
    .ColHidden(.Cols - 4) = True
    .ColHidden(.Cols - 3) = True
    .ColHidden(.Cols - 2) = True
    .ColHidden(.Cols - 1) = True
        
                
    .ExplorerBar = flexExSort
    
    .SubtotalPosition = flexSTBelow
    
    
    .Subtotal flexSTSum, -1, 6, "#", &HC0FFC0, , True, "≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, 8, "#", &HC0FFC0, , True, "≈Ã„«·Ì"
    
    
    If grid1.Rows > 1 Then
        aSub = AddFlag(Empty, "row", grid1.Rows - 1)
        aSub = AddFlag(aSub, "col", 0)
        aSub = AddFlag(aSub, "cols", 6)
        aSub = AddFlag(aSub, "text", "«·≈Ã„«·Ì")
        aRow = AddFlag(aRow, aSub)
    
        aSub = AddFlag(Empty, "row", grid1.Rows - 1)
        aSub = AddFlag(aSub, "col", 0)
        aSub = AddFlag(aSub, "cols", 6)
        aSub = AddFlag(aSub, "text", "«·≈Ã„«·Ì")
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "back_color", 19)
        aXl = AddFlag(aXl, aSub)
        
        
'        If grid1.ValueMatrix(1, 9) <> 0 Then
'            grid1.AddItem ""
'            grid1.AddItem ""
'
'            aSub = AddFlag(Empty, "row", grid1.Rows - 2)
'            aSub = AddFlag(aSub, "col", 0)
'            aSub = AddFlag(aSub, "cols", 8)
'            aRow = AddFlag(aRow, aSub)
'
'            aSub = AddFlag(Empty, "row", grid1.Rows - 2)
'            aSub = AddFlag(aSub, "col", 0)
'            aSub = AddFlag(aSub, "cols", 8)
'            aSub = AddFlag(aSub, "bold", True)
'            aSub = AddFlag(aSub, "back_color", 19)
'            aXl = AddFlag(aXl, aSub)
'
'            aSub = AddFlag(Empty, "row", grid1.Rows - 1)
'            aSub = AddFlag(aSub, "col", 0)
'            aSub = AddFlag(aSub, "cols", 8)
'            aRow = AddFlag(aRow, aSub)
'
'            aSub = AddFlag(Empty, "row", grid1.Rows - 1)
'            aSub = AddFlag(aSub, "col", 0)
'            aSub = AddFlag(aSub, "cols", 8)
'            aSub = AddFlag(aSub, "bold", True)
'            aSub = AddFlag(aSub, "back_color", 19)
'            aXl = AddFlag(aXl, aSub)
'
'
'            For i = 0 To 7
'                If i <= 5 Then grid1.TextMatrix(grid1.Rows - 3, i) = "≈Ã„«·Ì"
'                grid1.TextMatrix(grid1.Rows - 2, i) = "«·Œ’„"
'                grid1.TextMatrix(grid1.Rows - 1, i) = "«·’«›Ì"
'            Next
'
'            grid1.TextMatrix(grid1.Rows - 2, 8) = grid1.ValueMatrix(1, 9)
'            grid1.TextMatrix(grid1.Rows - 1, 8) = Round(grid1.ValueMatrix(grid1.Rows - 3, 8) - grid1.ValueMatrix(grid1.Rows - 2, 8), 2)
'
'            grid1.MergeCells = flexMergeRestrictRows
'
'            grid1.MergeRow(grid1.Rows - 3) = True
'            grid1.MergeRow(grid1.Rows - 2) = True
'            grid1.MergeRow(grid1.Rows - 1) = True
'
'            .Cell(flexcpBackColor, .Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = &HC0FFC0
'            .Cell(flexcpBackColor, .Rows - 2, 0, grid1.Rows - 1, grid1.Cols - 1) = &HC0FFC0
'
'
'        End If
    End If
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .Select .Rows - 1, 0
    .ShowCell .Rows - 1, 0
End With
End Sub
Private Sub myload()
Dim cString As String
cString = " SELECT FILE6_90.ITEM," & _
          " FILE1_10.MODELFACT0," & _
          " FACT.DESCA," & _
          " COALESCE(FILE1_10.desca,FILE6_90.ITEM_NAME)," & _
          " FILE1_10.COLOR," & _
          " FILE1_10.SCAL," & _
          " FILE6_90.QUANT," & _
          " FILE6_90.PRICE," & _
          " FILE6_90.TOTAL," & _
          " FILE6_90H.DISCOUNT," & _
          " FILE6_90H.DATE" & _
          " FROM  FILE6_90 " & _
          " INNER JOIN FILE6_90H ON FILE6_90.DOC_NO = FILE6_90H.DOC_NO " & _
          " LEFT  JOIN FILE1_10 ON FILE6_90.ITEM = FILE1_10.ITEM" & _
          " LEFT  JOIN  FACT ON FILE1_10.code = FACT.CODE" & _
          " WHERE FILE6_90H.DOC_NO =  " & MyParn(sDoc_no)
cString = cString & " ORDER BY DATE"

Dim con As New ADODB.Connection
On Error GoTo myerror
Set grid1.DataSource = myRs(cString)
If grid1.Rows > 1 Then
    xDoc_no.Caption = sDoc_no
    xdate.Caption = myFormat_p(grid1.TextMatrix(1, 10))
    aHeader(0) = "—ﬁ„ «·ÿ·»Ì… : " & xDoc_no.Caption
    aHeader(1) = "«· «—ÌŒ : " & xdate.Caption
End If
fixGrd
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set ShowOrderManfrm = Nothing
End Sub
Private Sub myLoadGrd()
If Not openCn(con) Then On Error GoTo Finally

If IsNull(rsValue("select order_no from vw_online_orders_open where order_no = " & MyParn(grid1.TextMatrix(grid1.Row, 0)))) Then
    Exit Sub
End If

Dim strSql As String
cString = " SELECT FILE6_90H_ER.DOC_NO," & _
          " FILE1_10.MODELFACT012," & _
          " FACT.DESCA," & _
          " COALESCE(FILE1_10.desca,FILE6_90.ITEM_NAME)," & _
          " FILE1_10.COLOR," & _
          " FILE1_10.SCAL," & _
          " FILE6_90.QUANT," & _
          " FILE6_90.PRICE," & _
          " FILE6_90.TOTAL," & _
          " FILE6_90H.DISCOUNT," & _
          " FILE6_90H.DATE" & _
          " FROM  FILE6_90 " & _
          " INNER JOIN FILE6_90H ON FILE6_90.DOC_NO = FILE6_90H.DOC_NO " & _
          " LEFT  JOIN FILE1_10 ON FILE6_90.ITEM = FILE1_10.ITEM" & _
          " LEFT  JOIN  FACT ON FILE1_10.code = FACT.CODE" & _
          " WHERE FILE6_90H.DOC_NO =  " & MyParn(sDoc_no)
cString = cString & " ORDER BY DATE"

Dim con As New ADODB.Connection


If openCn(con) Then
    On Error GoTo Finally
    Set grid1.DataSource = myRs(cString, con)
End If

If grid1.Rows > 1 Then
    xDoc_no.Caption = sDoc_no
    xdate.Caption = myFormat_p(grid1.TextMatrix(1, 10))
    aHeader(0) = "—ﬁ„ «·ÿ·»Ì… : " & xDoc_no.Caption
    aHeader(1) = "«· «—ÌŒ : " & xdate.Caption
End If
fixGrd
Finally:
closeCon con
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myloadGrdError()
Dim strSql As String
strSql = "SELECT STAGES_CODES.DESCA AS [«·„—Õ·…], " & _
         " MAN_ONLINE_CODES.DESCA AS [«·„‰œÊ»], " & _
         " FILE6_90H_ER.DESCA AS [«·»Ì«‰], " & _
         " FORMAT(FILE6_90H_ER.TIME, 'yyyy/M/d HH:mm') AS [«·Êﬁ ]" & _
         " FROM STAGES_CODES " & _
         "    INNER JOIN FILE6_90H_ER " & _
         "        ON STAGES_CODES.CODE = FILE6_90H_ER.STAGE " & _
         "    INNER JOIN MAN_ONLINE_CODES " & _
         "        ON FILE6_90H_ER.MAN = MAN_ONLINE_CODES.CODE" & _
         " WHERE FILE6_90H_ER.DOC_NO = " & MyParn(xDoc_no.Caption)

On Error GoTo myerror

Set grdError.DataSource = myRs(strSql)

grdError.ColWidth(0) = 1300
grdError.ColWidth(1) = 1500
grdError.ColWidth(2) = 5000
grdError.ColWidth(3) = 1600


Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

