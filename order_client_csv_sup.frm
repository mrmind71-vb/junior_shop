VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form orders_clients_csv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "„—«Ã⁄… «·‘Õ‰"
   ClientHeight    =   10590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10590
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9465
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   20265
      _ExtentX        =   35745
      _ExtentY        =   16695
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "«’‰«› »»«—ﬂÊœ €Ì— ’«·Õ"
      TabPicture(0)   =   "order_client_csv_sup.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "«’‰«› —’ÌœÂ« «ﬁ· „‰ «·ÿ·»Ì… «Ê »œÊ‰ —’Ìœ"
      TabPicture(1)   =   "order_client_csv_sup.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Option2(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Option2(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Option2(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "«’‰«› ’«·Õ… ··ÿ·»Ì…"
      TabPicture(2)   =   "order_client_csv_sup.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "grid1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Option1(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Option1(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Option1(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·ﬂ·"
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
         Height          =   285
         Index           =   0
         Left            =   -55965
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   9000
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "ÌÊÃœ —’Ìœ"
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
         Height          =   285
         Index           =   1
         Left            =   -59160
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   9000
         Width           =   1140
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "·« ÌÊÃœ —’Ìœ"
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
         Height          =   285
         Index           =   2
         Left            =   -62670
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   9000
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·—’Ìœ «ﬁ· „‰ «·ÿ·»Ì…"
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
         Height          =   285
         Index           =   2
         Left            =   11700
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   9000
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·ÿ·»Ì… «ﬁ· „‰ «Ê Ì”«ÊÌ —’Ìœ «·’‰› «Ê —’Ìœ €Ì— „ÿ·Ê»"
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
         Height          =   285
         Index           =   1
         Left            =   14085
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   9000
         Width           =   4920
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·ﬂ·"
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
         Height          =   285
         Index           =   0
         Left            =   18990
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   9000
         Width           =   1095
      End
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   8520
         Left            =   90
         TabIndex        =   9
         Top             =   360
         Width           =   20070
         _cx             =   35401
         _cy             =   15028
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   -2147483633
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
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
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   8520
         Left            =   -74910
         TabIndex        =   10
         Top             =   360
         Width           =   20085
         _cx             =   35428
         _cy             =   15028
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
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
      Begin VSFlex7Ctl.VSFlexGrid grid3 
         Height          =   8835
         Left            =   -74910
         TabIndex        =   11
         Top             =   360
         Width           =   20085
         _cx             =   35428
         _cy             =   15584
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
         BackColorFixed  =   6249205
         ForeColorFixed  =   0
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   -2147483633
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
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
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   690
      ScaleWidth      =   20400
      TabIndex        =   0
      Top             =   9900
      Width           =   20400
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   -45
         Width           =   8925
         Begin Threed.SSCommand cmdExcel 
            Height          =   510
            Left            =   3015
            TabIndex        =   2
            Top             =   135
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   900
            _Version        =   196610
            CaptionStyle    =   1
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
            Picture         =   "order_client_csv_sup.frx":0054
            Caption         =   " ’œÌ— Excel"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdPrint 
            Height          =   510
            Left            =   1440
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   135
            Width           =   1455
            _ExtentX        =   2566
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
            Picture         =   "order_client_csv_sup.frx":24BE
            Caption         =   " ÿ»«⁄…"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "order_client_csv_sup.frx":4834
         End
         Begin Threed.SSCommand cmdExit 
            Height          =   510
            Left            =   45
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
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
            Picture         =   "order_client_csv_sup.frx":69B7
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdCSV 
            Height          =   510
            Left            =   6750
            TabIndex        =   5
            Top             =   135
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   900
            _Version        =   196610
            CaptionStyle    =   1
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
            Picture         =   "order_client_csv_sup.frx":8CDA
            Caption         =   "”Õ» „‰ „·› CSV"
            TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
            ButtonStyle     =   3
            PictureAlignment=   9
         End
         Begin Threed.SSCommand cmdSave 
            Height          =   510
            Left            =   4995
            TabIndex        =   7
            Top             =   135
            Width           =   1725
            _ExtentX        =   3043
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
            Picture         =   "order_client_csv_sup.frx":B2AD
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            PictureDisabledFrames=   1
            ShapeSize       =   1
            PictureDisabled =   "order_client_csv_sup.frx":DBD2
         End
      End
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   9705
      Visible         =   0   'False
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "orders_clients_csv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
Public myform As Form
Private Sub cmdCSV_Click()
getData
End Sub

Private Sub cmdExcel_Click()
If SSTab1.Tab = 1 Then
    ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, TabStrip1.Tabs(1).Caption)
ElseIf SSTab1.Tab = 1 = 2 Then
    ToFileExelNew grid2, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, TabStrip1.Tabs(2).Caption)
ElseIf SSTab1.Tab = 1 = 3 Then
    ToFileExelNew grid3, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, TabStrip1.Tabs(3).Caption)
End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Set printGrdNew.myform = Me
If TabStrip1.SelectedItem.Index = 1 Then
    printGrdNew.doPrint grid1, 0.85, 0, "ÃÊ‰ÌÊ—", Me.Caption, ArbString(retHeader(aHeader, 0, 5)), TabStrip1.Tabs(1).Caption, False, True, 10
ElseIf TabStrip1.SelectedItem.Index = 2 Then
    printGrdNew.doPrint grid2, 0.85, 0, "ÃÊ‰ÌÊ—", Me.Caption, ArbString(retHeader(aHeader, 0, 5)), TabStrip1.Tabs(2).Caption, False, True, 10
ElseIf TabStrip1.SelectedItem.Index = 3 Then
    printGrdNew.doPrint grid3, 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption, ArbString(retHeader(aHeader, 0, 5)), TabStrip1.Tabs(2).Caption, False, False, 10
End If
printGrdNew.Show 1
End Sub
Private Function myReplace() As Boolean
With grid1
Dim sCaption As String
Dim i As Long
prog1.Visible = True
Dim db As New clsDb
Dim sb As New ChilkatStringBuilder
Dim nCount As Long
For i = 1 To .Rows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & grid1.Rows - 1
    prog1.Value = Round(i / (.Rows - 1), 2) * 100
    
    If nCount >= 1000 Then
        sb.Prepend "insert into file6_50(doc_no,ITEM,MODEL,QUANT,PRICE)" & _
                   "VALUES"
        db.Sql = sb.GetAsString
        sb.Clear
        nCount = 1
        Exit For
    Else
        nCount = nCount + 1
    End If
    
    If sb.length > 0 Then sb.Append ","
    sb.Append "("
    sb.Append addstring(sDoc_no) & ","
    sb.Append addstring(grid1.TextMatrix(i, 1)) & ","
    sb.Append addstring(grid1.TextMatrix(i, 3)) & ","
    sb.Append .ValueMatrix(i, 9) & ","
    sb.Append .ValueMatrix(i, 6)
    sb.Append ")"
Next

If sb.length > 0 Then
        sb.Prepend "insert into file6_50(doc_no,ITEM,MODEL,QUANT,PRICE)" & _
                   "VALUES"
    db.Sql = sb.GetAsString
End If

If Not db.ExecuteTransaction Then GoTo cleanUp

End With
prog1.Visible = False
myInform " „  ÕœÌÀ «·„” ‰œ«  »‰Ã«Õ"
bSave = True
myReplace = True
cleanUp:
prog1.Visible = True
Set db = Nothing
End Function

Private Sub cmdSave_Click()
If grid1.Rows <= 1 Then MsgBox "·«  ÊÃœ ”Ã·«  ··Õ›Ÿ"
If myReplace Then
    myform.myProc
    Unload Me
End If
End Sub

Private Sub Form_Load()
Fixgrd1
Fixgrd2
Fixgrd3
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set dbm = Nothing
Set online_ship_checkfrm = Nothing
End Sub
Private Sub grid3_DblClick()
If grid3.col = 1 And grid3.Row > 0 Then
    Clipboard.Clear
    Clipboard.SetText grid3.TextMatrix(grid3.Row, grid3.col)
    Inform " „ «·‰”Œ"
End If
End Sub
Private Function getData()
Dim cFileName As String
Dim sb As New ChilkatStringBuilder

Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

cFileName = Common1.FileName

If cFileName = "" Then Exit Function

Dim cSv As New ChilkatCsv
Dim i As Long

'MsgBox GetTextFileEncoding(cFileName)

nAccess = sb.LoadFile(cFileName, "auto")

If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If


cSv.HasColumnNames = 1
nAccess = cSv.LoadFromString(sb.GetAsString)

If cSv.NumRows < 1 Then Exit Function

myloadNew cSv
End Function
Private Function myloadNew(ByRef cSv As ChilkatCsv) As Boolean
Dim i As Long
Dim sCaption As String
Dim ship_no As String
Dim date_Pick As String
Dim cod As String
Dim Tb As New ChilkatStringBuilder


sCaption = Me.Caption
prog1.Visible = True

grid1.Rows = 1
grid1.Redraw = flexRDNone

grid2.Rows = 1
grid2.Redraw = flexRDNone

grid3.Rows = 1
grid3.Redraw = flexRDNone


Dim strSql As String
strSql = "SELECT * " & _
          " FROM vw_items_orders"
                       
Dim db As New clsDb
Dim locTable As ADODB.RecordSet
Set locTable = db.myRs(strSql)
If locTable Is Nothing Then GoTo cleanUp

For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
        
    sItem = Trim(cSv.GetCell(i, 0))
    nQuant = Trim(cSv.GetCell(i, 1))
                    
    If sItem <> "" And nQuant > 0 Then
         locTable.Find "ITEM = " & MyParn(sItem), , adSearchForward, adBookmarkFirst
         If Not locTable.EOF Then
            If Val(locTable!balance) > 0 Or Not locTable!check_Bal Then
                grid1.AddItem ""
                grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
                grid1.TextMatrix(grid1.Rows - 1, 1) = locTable!item & ""
                grid1.TextMatrix(grid1.Rows - 1, 2) = locTable!desca & ""
                grid1.TextMatrix(grid1.Rows - 1, 3) = locTable!model & ""
                grid1.TextMatrix(grid1.Rows - 1, 4) = locTable!color & ""
                grid1.TextMatrix(grid1.Rows - 1, 5) = locTable!scal & ""
                grid1.TextMatrix(grid1.Rows - 1, 6) = locTable!price & ""
                grid1.TextMatrix(grid1.Rows - 1, 7) = locTable!balance
                grid1.TextMatrix(grid1.Rows - 1, 8) = Round(nQuant / 12, 2)
                If Val(locTable!balance) >= Round(nQuant / 12) Or Not locTable!check_Bal Then
                    grid1.TextMatrix(grid1.Rows - 1, 9) = Round(nQuant / 12, 2)
                Else
                    grid1.TextMatrix(grid1.Rows - 1, 9) = Val(locTable!balance)
                End If
                grid1.TextMatrix(grid1.Rows - 1, 10) = IIf(locTable!check_Bal, 1, 0)
            End If
            If Val(locTable!balance) < (nQuant / 12) And locTable!check_Bal Then
                grid2.AddItem ""
                grid2.TextMatrix(grid2.Rows - 1, 0) = grid2.Rows - 1
                grid2.TextMatrix(grid2.Rows - 1, 1) = locTable!item & ""
                grid2.TextMatrix(grid2.Rows - 1, 2) = locTable!desca & ""
                grid2.TextMatrix(grid2.Rows - 1, 3) = locTable!model & ""
                grid2.TextMatrix(grid2.Rows - 1, 4) = locTable!color & ""
                grid2.TextMatrix(grid2.Rows - 1, 5) = locTable!scal & ""
                grid2.TextMatrix(grid2.Rows - 1, 6) = locTable!price & ""
                grid2.TextMatrix(grid2.Rows - 1, 7) = locTable!balance
                grid2.TextMatrix(grid2.Rows - 1, 8) = Round(nQuant / 12, 2)
                grid2.TextMatrix(grid2.Rows - 1, 9) = IIf(locTable!balance > 0, locTable!balance, 0)
            End If
        Else
            grid3.AddItem ""
            grid3.TextMatrix(grid3.Rows - 1, 0) = grid3.Rows - 1
            grid3.TextMatrix(grid3.Rows - 1, 1) = sItem
            grid2.TextMatrix(grid2.Rows - 1, 9) = nQuant
         End If
    End If
Next
cleanUp:
grid1.Redraw = flexRDBuffered
grid2.Redraw = flexRDBuffered
grid3.Redraw = flexRDBuffered
Fixgrd1
Fixgrd2
prog1.Visible = False
Me.Caption = sCaption
Set db = Nothing
End Function
Private Sub Fixgrd1()
With grid1

.Redraw = flexRDNone

.RowHeight(0) = 500
.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "«·»«—ﬂÊœ"
.TextMatrix(0, 2) = "«·»Ì«‰"
.TextMatrix(0, 3) = "«·„ÊœÌ·"
.TextMatrix(0, 4) = "«··Ê‰"
.TextMatrix(0, 5) = "«·„ﬁ«”"
.TextMatrix(0, 6) = "«·”⁄—"
.TextMatrix(0, 7) = "«·—’Ìœ"
.TextMatrix(0, 8) = "ﬂ„Ì… «·ÿ·»Ì…"
.TextMatrix(0, 9) = "ﬂ„Ì… „ƒﬂœ…"
.TextMatrix(0, 10) = "—’Ìœ „ÿ·Ê»"

.ColWidth(0) = 800
.ColWidth(1) = 2000
.ColWidth(2) = 6000
.ColWidth(3) = 1500
.ColWidth(4) = 2000
.ColWidth(5) = 1500
.ColWidth(6) = 1200
.ColWidth(7) = 1100
.ColWidth(8) = 1100
.ColWidth(9) = 1100
.ColWidth(10) = 1200


Dim i As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

For i = 1 To grid1.Rows - 1
    .TextMatrix(i, 0) = i
Next

.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDate
.ColDataType(8) = flexDTDate
.ColDataType(9) = flexDTDouble
.ColDataType(10) = flexDTBoolean
For i = 1 To .Rows - 1
    If Option1(1).Value Then
        .RowHidden(i) = .ValueMatrix(i, 7) < .ValueMatrix(i, 8)
    ElseIf Option1(2).Value Then
        .RowHidden(i) = .ValueMatrix(i, 7) >= .ValueMatrix(i, 8)
    Else
        .RowHidden(i) = False
    End If
Next
.Redraw = flexRDBuffered
End With
End Sub
Private Sub Fixgrd2()
With grid2
.RowHeight(0) = 500

.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "«·»«—ﬂÊœ"
.TextMatrix(0, 2) = "«·»Ì«‰"
.TextMatrix(0, 3) = "«·„ÊœÌ·"
.TextMatrix(0, 4) = "«··Ê‰"
.TextMatrix(0, 5) = "«·„ﬁ«”"
.TextMatrix(0, 6) = "«·”⁄—"
.TextMatrix(0, 7) = "«·—’Ìœ"
.TextMatrix(0, 8) = "ﬂ„Ì… «·ÿ·»Ì…"
.TextMatrix(0, 9) = "ﬂ„Ì… „ƒﬂœ…"

.ColWidth(0) = 800
.ColWidth(1) = 2000
.ColWidth(2) = 6000
.ColWidth(3) = 1500
.ColWidth(4) = 2000
.ColWidth(5) = 2000
.ColWidth(6) = 1200
.ColWidth(7) = 1100
.ColWidth(8) = 1100
.ColWidth(9) = 1100

Dim i As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
Next

.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDate
.ColDataType(8) = flexDTDate
.ColDataType(9) = flexDTDouble

For i = 1 To .Rows - 1
    If Option2(1).Value Then
        .RowHidden(i) = .ValueMatrix(i, 7) <= 0
    ElseIf Option2(2).Value Then
        .RowHidden(i) = .ValueMatrix(i, 7) > 0
    Else
        .RowHidden(i) = False
    End If
Next

End With
End Sub
Private Sub Fixgrd3()
With grid3
.RowHeight(0) = 800
.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "«·»«—ﬂÊœ"
.TextMatrix(0, 2) = "«·»Ì«‰"
.TextMatrix(0, 3) = "«·„ÊœÌ·"
.TextMatrix(0, 4) = "«··Ê‰"
.TextMatrix(0, 5) = "«·„ﬁ«”"
.TextMatrix(0, 6) = "«·”⁄—"
.TextMatrix(0, 7) = "«·—’Ìœ"
.TextMatrix(0, 8) = "ﬂ„Ì… «·ÿ·»Ì…"
.TextMatrix(0, 9) = "ﬂ„Ì… „ƒﬂœ…"

Dim i As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
Next

.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDate
.ColDataType(8) = flexDTDate
.ColDataType(9) = flexDTDouble
End With
End Sub
Private Sub Option1_Click(Index As Integer)
Fixgrd1
End Sub
Private Sub Option2_Click(Index As Integer)
Fixgrd2
End Sub
