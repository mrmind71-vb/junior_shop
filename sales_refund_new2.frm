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
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   1095
      ScaleWidth      =   20370
      TabIndex        =   23
      Top             =   9105
      Width           =   20370
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   1005
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   45
         Width           =   10905
         Begin VB.Label xdiscount_offer_ret_Rate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   225
            Width           =   600
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "’«›Ì"
            Height          =   285
            Left            =   1395
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   270
            Width           =   555
         End
         Begin VB.Label xDiscount_offer_ret 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3015
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   225
            Width           =   915
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Œ’„ ⁄—Ê÷"
            Height          =   285
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   225
            Width           =   1050
         End
         Begin VB.Label xtotal_item_ret 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5355
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   225
            Width           =   1455
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "≈Ã„«·Ì"
            Height          =   285
            Left            =   6885
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   225
            Width           =   600
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ﬂ„Ì… „— Ã⁄"
            Height          =   285
            Left            =   9765
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   225
            Width           =   870
         End
         Begin VB.Label xQuant_Ret 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   8235
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   225
            Width           =   1455
         End
         Begin VB.Label xtotal_offer_ret 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   225
            Width           =   1230
         End
         Begin VB.Label xdiscount_add_Ret 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   8730
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   585
            Width           =   960
         End
         Begin VB.Label Label23 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Œ’„ «÷«›Ì"
            Height          =   285
            Left            =   9765
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   630
            Width           =   1005
         End
         Begin VB.Label xdiscount_add_ret_Rate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   8235
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   585
            Width           =   465
         End
         Begin VB.Label Label25 
            BackColor       =   &H00FFFFFF&
            Caption         =   "≈Ã„«·Ì «·Œ’„"
            Height          =   285
            Left            =   6885
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   630
            Width           =   1185
         End
         Begin VB.Label xdiscount_ret 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5850
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   585
            Width           =   960
         End
         Begin VB.Label xdiscount_Ret_rate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5355
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   585
            Width           =   465
         End
         Begin VB.Label Label28 
            BackColor       =   &H00FFFFFF&
            Caption         =   "’«›Ì «·›« Ê—…"
            Height          =   285
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   630
            Width           =   1095
         End
         Begin VB.Label xtotal_Ret 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   585
            Width           =   1545
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1050
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   0
         Width           =   4605
         Begin Threed.SSCommand cmdExit 
            Height          =   825
            Left            =   45
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1455
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
            Picture         =   "sales_refund_new2.frx":0000
            Caption         =   "Œ—ÊÃ"
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdSelectAll 
            Height          =   825
            Left            =   2205
            TabIndex        =   26
            Top             =   180
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   1455
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
            Picture         =   "sales_refund_new2.frx":2323
            Caption         =   "«Œ Ì«— «·ﬂ·"
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            PictureDisabledFrames=   1
            ShapeSize       =   1
            PictureDisabled =   "sales_refund_new2.frx":473F
         End
         Begin Threed.SSCommand cmdSend 
            Height          =   825
            Left            =   3375
            TabIndex        =   52
            Top             =   180
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   1455
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
            Picture         =   "sales_refund_new2.frx":6A4A
            Caption         =   "Õ›Ÿ"
            Alignment       =   8
            ButtonStyle     =   2
            PictureAlignment=   11
            BevelWidth      =   0
            PictureDisabledFrames=   1
            ShapeSize       =   1
            PictureDisabled =   "sales_refund_new2.frx":936F
         End
         Begin Threed.SSCommand cmddelAll 
            Height          =   825
            Left            =   1170
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   180
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1455
            _Version        =   196610
            ForeColor       =   255
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
            Picture         =   "sales_refund_new2.frx":BBC3
            Caption         =   "Õ–› «·ﬂ·"
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            PictureDisabledFrames=   1
            ShapeSize       =   1
            PictureDisabled =   "sales_refund_new2.frx":E1A5
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   6180
      Begin VB.TextBox xItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   2265
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·’‰›"
         Height          =   285
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   645
      End
      Begin VB.Label xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   585
         Width           =   5100
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·»«—ﬂÊœ"
         Height          =   285
         Left            =   5265
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   6255
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   -45
      Width           =   9240
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·›—⁄"
         Height          =   285
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   225
         Width           =   555
      End
      Begin VB.Label xBranch 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   225
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·›« Ê—…"
         Height          =   330
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·›« Ê—…"
         Height          =   330
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   225
         Width           =   645
      End
      Begin VB.Label xdoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   225
         Width           =   1680
      End
      Begin VB.Label xdiscount_rate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   945
         Width           =   555
      End
      Begin VB.Label xdiscount_Add_rate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   945
         Width           =   510
      End
      Begin VB.Label xDiscount_offer_Rate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   585
         Width           =   555
      End
      Begin VB.Label xDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì «·Œ’„"
         Height          =   285
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "’«›Ì «·›« Ê—…"
         Height          =   285
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Œ’„ «÷«›Ì"
         Height          =   285
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label xTotal_offer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   585
         Width           =   1725
      End
      Begin VB.Label xDiscount_offer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   585
         Width           =   1230
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Œ’„ «·⁄—Ê÷"
         Height          =   330
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·≈Ã„«·Ì"
         Height          =   330
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   585
         Width           =   690
      End
      Begin VB.Label xTotal_item 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   1680
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·’«›Ì"
         Height          =   330
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   735
      End
      Begin VB.Label xtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   945
         Width           =   1725
      End
      Begin VB.Label xDiscount_add 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6885
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   945
         Width           =   1140
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3795
      Left            =   45
      TabIndex        =   0
      Top             =   1350
      Width           =   15450
      _cx             =   27252
      _cy             =   6694
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
      Cols            =   18
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
      Height          =   3930
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5175
      Width           =   15495
      _cx             =   27331
      _cy             =   6932
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
   Begin VB.Label xPhone 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "sales_refundfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
Dim loctable As New ADODB.Recordset
Public sOnline_doc As String
Public sFlag As String
Public myForm As Form
Public bOnline As Boolean
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
grid1.Rows = 1
CalcTotals
addRow
fixGrdColor
End Sub
Private Sub cmdsend_Click()
myForm.myReplaceRefund -1 * Val(xdiscount_add_Ret.Caption), -1 * Val(xdiscount_ret.Caption)
End Sub

Private Sub Form_Load()
openCon con

Set grid1.DataSource = data1
Set grid2.DataSource = data2

myLoadInv

myloadgrd2

Fixgrd
CalcTotals

CellPos 13, 0, grid1.Cols - 1
End Sub
Private Sub myLoadInv()
Dim cString As String
cString = "SELECT FILE6_20H.DATE," & _
          "FILE6_20H.TOTAL_ITEM," & _
          "FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT AS TOTAL," & _
          "FILE6_20H.DISCOUNT_OFFER," & _
          "FILE6_20H.DISCOUNT_OFFER_RATE," & _
          "FILE6_20H.DISCOUNT_ADD," & _
          "FILE6_20H.DISCOUNT_ADD_RATE," & _
          "FILE6_20H.DISCOUNT," & _
          "FILE6_20H.DISCOUNT_RATE," & _
          "FILE0_40.DESCA AS BRANCH," & _
          "FILE6_20H.PHONE" & _
          " FROM FILE6_20H " & _
          " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE" & _
          " WHERE FILE6_20H.DOC_NO = " & MyParn(sDoc_no)
If Not bOnline Then

cString = cString & _
          " UNION ALL " & _
          "SELECT FR6_20H.DATE," & _
          "FR6_20H.TOTAL_ITEM," & _
          "FR6_20H.TOTAL_ITEM - FR6_20H.DISCOUNT AS TOTAL," & _
          "Fr6_20H.DISCOUNT_OFFER," & _
          "FR6_20H.DISCOUNT_OFFER_RATE," & _
          "FR6_20H.DISCOUNT_ADD," & _
          "FR6_20H.DISCOUNT_ADD_RATE," & _
          "FR6_20H.DISCOUNT," & _
          "FR6_20H.DISCOUNT_RATE," & _
          "BRANCH_FR.DESCA AS BRANCH," & _
          "FR6_20H.PHONE" & _
          " FROM FR6_20H " & _
          " INNER JOIN BRANCH_FR ON FR6_20H.STORE = BRANCH_FR.CODE" & _
          " WHERE FR6_20H.DOC_NO = " & MyParn(sDoc_no)
End If
          
Set loctable = cmd(cString, con).Execute
If Not loctable.EOF Then
    xdoc_no.Caption = sDoc_no
    xBranch.Caption = loctable!branch & ""
    If IsValidMobile(loctable!phone & "") Then
        xPhone.Caption = loctable!phone
    End If
    xdate.Caption = myFormat_p(loctable!Date)
        
    xTotal_item.Caption = Myvalue(loctable!TOTAL_ITEM)
    xDiscount_offer.Caption = Myvalue(loctable!DISCOUNT_OFFER)
    xDiscount_offer_Rate.Caption = Format(loctable!discount_offer_rate, "0%")
    xTotal_offer.Caption = loctable!TOTAL_ITEM - loctable!DISCOUNT_OFFER
    xDiscount_add.Caption = Myvalue(loctable!discount_add)
    xdiscount_Add_rate.Caption = Format(loctable!discount_add_Rate, "0%")
    xdiscount_Add_rate.Tag = loctable!discount_add_Rate
    xDiscount.Caption = Myvalue(loctable!DISCOUNT)
    xdiscount_rate.Caption = Format(loctable!discount_Rate, "0%")
    xdiscount_rate.Tag = loctable!discount_Rate
    
    xtotal.Caption = Myvalue(loctable!TOTAL)
End If
End Sub
Private Sub myloadgrd2()
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DOC_NO", sDoc_no)
If Not bOnline Then
    Set data2.Recordset = mycmd("dbo.sp_SALES_REFUND", con, adStoredProc, aPrm)
Else
    Set data2.Recordset = mycmd("dbo.sp_SALES_REFUND_ONLINE", con, adStoredProc, aPrm)
End If
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
    .TextMatrix(0, 18 + 1) = "ﬂÊœ «·»«∆⁄"
    .TextMatrix(0, 19 + 1) = "»Ì«‰ «·⁄—÷"
    
    .TextMatrix(0, 20 + 1) = "Flag"
    .TextMatrix(0, 21 + 1) = "„— Ã⁄"
    
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
    .ColWidth(19 + 1) = 1600
    
    .ColHidden(11) = True
    .ColHidden(13) = True
    .ColHidden(14) = True
    .ColHidden(15) = True
    .ColHidden(16) = True
    .ColHidden(17) = True
    .ColHidden(18 + 1) = True
    
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
    .MergeCol(19 + 1) = True
End With
End Sub
Private Sub CalcTotals()
grid1.Subtotal flexSTClear
grid2.Subtotal flexSTClear

cmddelAll.Enabled = grid1.Rows > 1
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
        xQuant_Ret.Caption = grid1.ValueMatrix(grid1.Rows - 1, 8)
        xtotal_item_ret.Caption = xTotal_item.Caption
        xDiscount_offer_ret.Caption = xDiscount_offer.Caption
        xdiscount_offer_ret_Rate.Caption = xDiscount_offer_Rate.Caption
        xtotal_offer_ret.Caption = xTotal_offer.Caption
        xdiscount_add_Ret.Caption = xDiscount_add.Caption
        xdiscount_add_ret_Rate.Caption = xdiscount_Add_rate.Caption
        xdiscount_ret.Caption = xDiscount.Caption
        xdiscount_Ret_rate.Caption = xdiscount_rate.Caption
        xtotal_Ret.Caption = xtotal.Caption
    Else
        For i = 1 To .Rows - 2
            If .ValueMatrix(i, 11) <> 0 Then
                If nOffer <> .ValueMatrix(i, 11) Then
                    nOffer = .ValueMatrix(i, 11)
                    nDiscountOffer = nDiscountOffer + .ValueMatrix(i, 13)
                End If
            End If
        Next
        
        xQuant_Ret.Caption = grid1.ValueMatrix(grid1.Rows - 1, 8)
        xtotal_item_ret.Caption = grid1.ValueMatrix(grid1.Rows - 1, 10)
    
        xDiscount_offer_ret.Caption = nDiscountOffer
        If grid1.ValueMatrix(grid1.Rows - 1, 10) <> 0 Then
            xdiscount_offer_ret_Rate.Caption = Format(nDiscountOffer / grid1.ValueMatrix(grid1.Rows - 1, 10), "0%")
        Else
            xdiscount_offer_ret_Rate.Caption = ""
        End If
        
        
        xtotal_offer_ret.Caption = grid1.ValueMatrix(grid1.Rows - 1, 10) - nDiscountOffer
        xdiscount_add_Ret.Caption = mRound(Val(xdiscount_Add_rate.Tag) * Val(xtotal_offer_ret.Caption))
        If Val(xtotal_offer_ret.Caption) <> 0 Then
            xdiscount_add_ret_Rate.Caption = Format(xdiscount_add_Ret.Caption / Val(xtotal_offer_ret.Caption), "0%")
        Else
            xdiscount_add_ret_Rate.Caption = ""
        End If
                                
        xdiscount_ret.Caption = Val(xDiscount_offer_ret.Caption) + Val(xdiscount_add_Ret.Caption)
    
        If Val(xDiscount_offer.Caption) = 0 And Val(xDiscount_add.Caption) = 0 And Val(xdiscount_rate.Tag) <> 0 Then
            xdiscount_ret.Caption = mRound(Val(xdiscount_rate.Tag) * Val(xtotal_item_ret.Caption))
        End If

        
        If grid1.ValueMatrix(grid1.Rows - 1, 10) <> 0 Then
            xdiscount_Ret_rate.Caption = Format(Val(xdiscount_ret.Caption) / grid1.ValueMatrix(grid1.Rows - 1, 10), "0%")
        Else
            xdiscount_Ret_rate.Caption = ""
        End If
        xtotal_Ret.Caption = grid1.ValueMatrix(grid1.Rows - 1, 10) - Val(xdiscount_ret.Caption)
    End If
Else
    xQuant_Ret.Caption = ""
    xtotal_item_ret.Caption = ""
    xDiscount_offer_ret.Caption = ""
    xdiscount_offer_ret_Rate.Caption = ""
    xdiscount_add_Ret.Caption = ""
    xdiscount_add_ret_Rate.Caption = ""
    xdiscount_ret.Caption = ""
    xdiscount_Ret_rate.Caption = ""
End If
End With
End Sub
Sub Fixgrd()
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
    .TextMatrix(0, 15 + 1) = "«·»«∆⁄"
    .TextMatrix(0, 16 + 1) = "Flag"
        
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
    .ColHidden(15 + 1) = True
    .ColHidden(16 + 1) = True
    
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
Private Sub Form_Resize()
Dim nHeight As Long
nHeight = Me.Height - (grid2.Top + 1000)
If nHeight > 3000 Then
    grid2.Height = nHeight
End If
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
Private Function myValid() As Boolean
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
myValid = True
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
        grid1.TextMatrix(grid1.Rows - 1, 14) = grid2.TextMatrix(Row, 19 + 1)
        grid1.TextMatrix(grid1.Rows - 1, 15 + 1) = grid2.TextMatrix(Row, 18 + 1)
        grid1.TextMatrix(grid1.Rows - 1, grid1.Cols - 1) = grid2.TextMatrix(Row, grid2.Cols - 2)
    End If
Next
CalcTotals
fixSerial
End With
End Sub

Private Sub xdiscount_add_rate_ret_Click()

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

    If IsNull(cm.Parameters("@ITEM").Value) Then
        sMsg = "ﬂÊœ €Ì— ’«·Õ"
        Exit Sub
    End If

    If cm.Parameters("@ITEM").Value <> xItem.text Then
        xItem.text = cm.Parameters("@ITEM").Value
    End If
    
    xdesca.Caption = cm.Parameters("@DESCA").Value
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

