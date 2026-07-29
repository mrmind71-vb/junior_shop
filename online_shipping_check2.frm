VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form online_ship_checkfrm 
   Caption         =   "„—«Ã⁄… «·‘Õ‰"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19815
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   19815
   StartUpPosition =   2  'CenterScreen
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
      ScaleWidth      =   19815
      TabIndex        =   4
      Top             =   7575
      Width           =   19815
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   -45
         Width           =   9105
         Begin Threed.SSCommand cmdExcel 
            Height          =   510
            Left            =   3195
            TabIndex        =   6
            Top             =   135
            Width           =   1590
            _ExtentX        =   2805
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
            Picture         =   "online_shipping_check2.frx":0000
            Caption         =   " ’œÌ— Excel"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdPrint 
            Height          =   510
            Left            =   1530
            TabIndex        =   7
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
            Picture         =   "online_shipping_check2.frx":246A
            Caption         =   " ÿ»«⁄…"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "online_shipping_check2.frx":47E0
         End
         Begin Threed.SSCommand cmdExit 
            Height          =   510
            Left            =   45
            TabIndex        =   8
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
            Picture         =   "online_shipping_check2.frx":6963
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdCSV 
            Height          =   510
            Left            =   4995
            TabIndex        =   9
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
            Picture         =   "online_shipping_check2.frx":8C86
            Caption         =   "”Õ» „‰ „·› CSV"
            TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
            ButtonStyle     =   3
            PictureAlignment=   9
         End
         Begin Threed.SSCommand cmdSave 
            Height          =   510
            Left            =   7335
            TabIndex        =   11
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
            Picture         =   "online_shipping_check2.frx":B259
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            PictureDisabledFrames=   1
            ShapeSize       =   1
            PictureDisabled =   "online_shipping_check2.frx":DB7E
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4245
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "»Ê«·’ ‘Õ‰ „ÿ«»ﬁ… "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "»Ê«·’ ‘Õ‰ „ÿ«»ﬁ… Ê „ ”Õ»Â« „‰ ﬁ»·"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "»Ê«·’ ‘Õ‰ €Ì— „”Ã·…"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   5505
      Left            =   405
      TabIndex        =   1
      Top             =   2700
      Width           =   18735
      _cx             =   33046
      _cy             =   9710
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
      Cols            =   15
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5505
      Left            =   -360
      TabIndex        =   2
      Top             =   1170
      Width           =   18735
      _cx             =   33046
      _cy             =   9710
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
      Cols            =   15
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
      Height          =   5505
      Left            =   0
      TabIndex        =   3
      Top             =   405
      Width           =   18735
      _cx             =   33046
      _cy             =   9710
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
      Cols            =   8
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
      TabIndex        =   10
      Top             =   7380
      Visible         =   0   'False
      Width           =   19815
      _ExtentX        =   34951
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "online_ship_checkfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCSV_Click()
getData
End Sub

Private Sub cmdExcel_Click()
If TabStrip1.SelectedItem.index = 1 Then
    ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, TabStrip1.Tabs(1).Caption)
ElseIf TabStrip1.SelectedItem.index = 2 Then
    ToFileExelNew grid2, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, TabStrip1.Tabs(2).Caption)
ElseIf TabStrip1.SelectedItem.index = 3 Then
    ToFileExelNew grid3, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, TabStrip1.Tabs(3).Caption)
End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Set printGrdNew.myForm = Me
If TabStrip1.SelectedItem.index = 1 Then
    printGrdNew.doprint grid1, 0.85, 0, "ÃÊ‰ÌÊ—", Me.Caption, ArbString(retHeader(aHeader, 0, 5)), TabStrip1.Tabs(1).Caption, False, True, 10
ElseIf TabStrip1.SelectedItem.index = 2 Then
    printGrdNew.doprint grid2, 0.85, 0, "ÃÊ‰ÌÊ—", Me.Caption, ArbString(retHeader(aHeader, 0, 5)), TabStrip1.Tabs(2).Caption, False, True, 10
ElseIf TabStrip1.SelectedItem.index = 3 Then
    printGrdNew.doprint grid3, 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption, ArbString(retHeader(aHeader, 0, 5)), TabStrip1.Tabs(2).Caption, False, False, 10
End If
printGrdNew.Show 1
End Sub

Private Sub cmdSave_Click()
If TabStrip1.SelectedItem.index = 1 Then
    If MsgBox("«÷«›…  «—ÌŒ «·‘Õ‰ ", vbOKCancel) <> vbOK Then Exit Sub
    If grid1.Rows = 1 Then
        MsgBox ("·«  ÊÃœ ”Ã·«  ··Õ›Ÿ")
        Exit Sub
    End If
    myReplace1
ElseIf TabStrip1.SelectedItem.index = 2 Then
    If MsgBox(" ÕœÌÀ  «—ÌŒ «·‘Õ‰ ", vbOKCancel) <> vbOK Then Exit Sub
    If grid2.Rows = 1 Then
        MsgBox ("·«  ÊÃœ ”Ã·«  ··Õ›Ÿ")
        Exit Sub
    End If
    myReplace2
End If
End Sub
Private Function myReplace1() As Boolean
With grid1
Dim sCaption As String
Dim i As Long
prog1.Visible = True
Dim db As New clsDb
For i = 1 To .Rows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & grid1.Rows - 1
    prog1.Value = Round(i / (.Rows - 1), 2) * 100
    If db.Execute("UPDATE FILE6_90S " & _
                 " SET DATE_PICK = " & addDate(.TextMatrix(i, 10)) & "," & _
                 " COD = " & .ValueMatrix(i, 11) & "," & _
                 " TYPE = " & addstring(.TextMatrix(i, 13)) & _
                 " FROM FILE6_90s " & _
                 " WHERE ID = " & .TextMatrix(i, .Cols - 1)) = -1 Then Exit Function
Next
End With
prog1.Visible = False
Inform " „  ÕœÌÀ «·„” ‰œ«  »‰Ã«Õ"
Set db = Nothing
End Function
Private Function myReplace2() As Boolean
Dim sCaption As String
Dim i As Long
With grid2
prog1.Visible = True
Dim db As New clsDb
For i = 1 To .Rows - 1
    prog1.Value = Round(i / (.Rows - 1), 2) * 100
    If db.Execute("UPDATE FILE6_90S " & _
                 " SET DATE_PICK = " & addDate(.TextMatrix(i, 10)) & "," & _
                 " COD = " & .ValueMatrix(i, 11) & "," & _
                 " TYPE = " & addstring(.TextMatrix(i, 13)) & _
                 " FROM FILE6_90s " & _
                 " WHERE ID = " & .TextMatrix(i, .Cols - 1)) = -1 Then Exit Function
Next
End With
prog1.Visible = False
Inform " „  ÕœÌÀ «·„” ‰œ«  »‰Ã«Õ"
Set db = Nothing
End Function
Private Sub Form_Load()
grid1.Visible = True
grid2.Visible = False
grid3.Visible = False

Fixgrd1
Fixgrd2
Fixgrd3
End Sub
Private Sub Form_Resize()
    ' ????? ????? ?? ??????? ?? ??? ?? ????? ??????? ???????
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' ??? ??? TabStrip ???? ???? ????? ???????
    TabStrip1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 900
    
    ' ??? ??? ?????? ????? ????? ??????? ???????? ??? TabStrip
    grid1.Move TabStrip1.ClientLeft, _
               TabStrip1.ClientTop, _
               TabStrip1.ClientWidth, _
               TabStrip1.ClientHeight
               
    ' ??? ??? ?????? ?????? ???? ????????
    grid2.Move TabStrip1.ClientLeft, _
               TabStrip1.ClientTop, _
               TabStrip1.ClientWidth, _
               TabStrip1.ClientHeight

    grid3.Move TabStrip1.ClientLeft, _
               TabStrip1.ClientTop, _
               TabStrip1.ClientWidth, _
               TabStrip1.ClientHeight

    grid1.ZOrder 0
    grid2.ZOrder 0
    grid3.ZOrder 0
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

Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.index = 1 Then
        grid1.Visible = True
        grid2.Visible = False
        grid3.Visible = False
    ElseIf TabStrip1.SelectedItem.index = 2 Then
        grid1.Visible = False
        grid2.Visible = True
        grid3.Visible = False
    ElseIf TabStrip1.SelectedItem.index = 3 Then
        grid1.Visible = False
        grid2.Visible = False
        grid3.Visible = True
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
Private Function myload(ByRef cSv As ChilkatCsv) As Boolean
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

    For i = 0 To cSv.NumRows - 1
        Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
        prog1.Value = Round(i / (cSv.NumRows), 2) * 100
            
        ship_no = Trim(cSv.GetCellByName(i, "Tracking Number"))
        
        date_Pick = Trim(cSv.GetCellByName(i, "Picked-Up Date"))
        date_Pick = Replace(date_Pick, ",", "")
        
        If IsDate(date_Pick) Then
            date_Pick = Format(date_Pick, "mm-dd-yyyy hh:nn")
            date_Pick = Format(date_Pick, "yyyy-mm-dd hh:nn")
        Else
            date_Pick = ""
        End If
        
        cod = cSv.GetCellByName(i, "Cod Amount")
        
        If ship_no <> "" And date_Pick <> "" Then
            Dim strSql As String
            strSql = "SELECT " & _
                      "s.SHIP_NO," & _
                      "s.ORDER_NO," & _
                      "s.ORDER_NO_SUP," & _
                      "v.TYPE_DESCA," & _
                      "v.NAME," & _
                      "v.PHONE," & _
                      "((v.TOTAL_ITEM - v.DISCOUNT) + v.charge2) as total," & _
                      "s.DOC_NO," & _
                      "H.DATE," & _
                      "s.date_pick," & _
                      "v.payment_method," & _
                      "s.ID" & _
                      " FROM FILE6_90S s" & _
                      " INNER JOIN FILE6_90SH h on s.DOC_NO = h.doc_no" & _
                      " INNER JOIN vw_online_orders v ON s.order_no = v.order_no AND s.ORDER_NO_SUP = v.ORDER_NO_SUP" & _
                      " WHERE s.SHIP_NO = " & MyParn(ship_no)
             
             Set locTable = db.myRs(strSql)
                      
             If locTable Is Nothing Then GoTo cleanUp
                      
             If Not locTable.EOF Then
                If IsNull(locTable!date_Pick) Then
                    grid1.AddItem ""
                    grid1.TextMatrix(grid1.Rows - 1, 0) = i
                    grid1.TextMatrix(grid1.Rows - 1, 1) = locTable!ship_no & ""
                    grid1.TextMatrix(grid1.Rows - 1, 2) = locTable!ORDER_NO & ""
                    grid1.TextMatrix(grid1.Rows - 1, 3) = locTable!order_no_sup & ""
                    grid1.TextMatrix(grid1.Rows - 1, 4) = locTable!TYPE_dESCA & ""
                    grid1.TextMatrix(grid1.Rows - 1, 5) = locTable!Name & ""
                    grid1.TextMatrix(grid1.Rows - 1, 6) = locTable!Phone & ""
                    grid1.TextMatrix(grid1.Rows - 1, 7) = locTable!TOTAL & ""
                    grid1.TextMatrix(grid1.Rows - 1, 8) = locTable!doc_no & ""
                    grid1.TextMatrix(grid1.Rows - 1, 9) = locTable!Date & ""
                    grid1.TextMatrix(grid1.Rows - 1, 10) = date_Pick
                    grid1.TextMatrix(grid1.Rows - 1, 11) = cod
                    grid1.TextMatrix(grid1.Rows - 1, 12) = locTable!Payment_Method & ""
                    grid1.TextMatrix(grid1.Rows - 1, 13) = locTable!ID
                Else
                    grid2.AddItem ""
                    grid2.TextMatrix(grid2.Rows - 1, 0) = i
                    grid2.TextMatrix(grid2.Rows - 1, 1) = locTable!ship_no & ""
                    grid2.TextMatrix(grid2.Rows - 1, 2) = locTable!ORDER_NO & ""
                    grid2.TextMatrix(grid2.Rows - 1, 3) = locTable!order_no_sup & ""
                    grid2.TextMatrix(grid2.Rows - 1, 4) = locTable!TYPE_dESCA & ""
                    grid2.TextMatrix(grid2.Rows - 1, 5) = locTable!Name & ""
                    grid2.TextMatrix(grid2.Rows - 1, 6) = locTable!Phone & ""
                    grid2.TextMatrix(grid2.Rows - 1, 7) = locTable!TOTAL & ""
                    grid2.TextMatrix(grid2.Rows - 1, 8) = locTable!doc_no & ""
                    grid2.TextMatrix(grid2.Rows - 1, 9) = locTable!Date & ""
                    grid2.TextMatrix(grid2.Rows - 1, 10) = date_Pick
                    grid2.TextMatrix(grid2.Rows - 1, 11) = cod
                    grid2.TextMatrix(grid2.Rows - 1, 12) = locTable!Payment_Method & ""
                    grid2.TextMatrix(grid2.Rows - 1, 13) = locTable!ID
                End If
             End If
        End If
    Next
cleanUp:
grid1.Redraw = flexRDBuffered
grid2.Redraw = flexRDBuffered
prog1.Visible = False
Me.Caption = sCaption
Set db = Nothing
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

Dim sb As New ChilkatStringBuilder
For i = 0 To cSv.NumRows - 1
    If Trim(cSv.GetCellByName(i, "Tracking Number")) <> "" Then
        sb.Append MyParn(cSv.GetCellByName(i, "Tracking Number")) & ","
    End If
Next

sb.Shorten 1
If sb.length = 0 Then Exit Function


Dim strSql As String
strSql = "SELECT " & _
          "s.SHIP_NO," & _
          "s.ORDER_NO," & _
          "s.ORDER_NO_SUP," & _
          "v.TYPE_DESCA," & _
          "v.NAME," & _
          "v.PHONE," & _
          "((v.TOTAL_ITEM - v.DISCOUNT) + v.charge2) as total," & _
          "s.DOC_NO," & _
          "H.DATE," & _
          "s.date_pick," & _
          "v.payment_method," & _
          "s.ID" & _
          " FROM FILE6_90S s" & _
          " INNER JOIN FILE6_90SH h on s.DOC_NO = h.doc_no" & _
          " INNER JOIN vw_online_orders v ON s.order_no = v.order_no AND s.ORDER_NO_SUP = v.ORDER_NO_SUP" & _
          " WHERE s.SHIP_NO in (" & sb.GetAsString & ")"
                       
Dim db As New clsDb
Dim locTable As ADODB.Recordset
Set locTable = db.myRs(strSql)
If locTable Is Nothing Then GoTo cleanUp

For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
        
    ship_no = Trim(cSv.GetCellByName(i, "Tracking Number"))
    
    date_Pick = Trim(cSv.GetCellByName(i, "Picked-Up Date"))
    date_Pick = Replace(date_Pick, ",", "")
    
    ref = cSv.GetCellByName(i, "Business Reference Number")
    sType = cSv.GetCellByName(i, "Type")
    sName = cSv.GetCellByName(i, "Consignee Name")
    Phone = cSv.GetCellByName(i, "Consignee phone")
    
    If IsDate(date_Pick) Then
        date_Pick = Format(date_Pick, "mm-dd-yyyy hh:nn")
        date_Pick = Format(date_Pick, "yyyy-mm-dd hh:nn")
    Else
        date_Pick = ""
    End If
    
    cod = cSv.GetCellByName(i, "Cod Amount")
            
    If ship_no <> "" And date_Pick <> "" Then
         locTable.Find "SHIP_NO = " & MyParn(ship_no), , adSearchForward, adBookmarkFirst
         If Not locTable.EOF Then
            If IsNull(locTable!date_Pick) Then
                grid1.AddItem ""
                grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
                grid1.TextMatrix(grid1.Rows - 1, 1) = locTable!ship_no & ""
                grid1.TextMatrix(grid1.Rows - 1, 2) = locTable!ORDER_NO & ""
                grid1.TextMatrix(grid1.Rows - 1, 3) = locTable!order_no_sup & ""
                grid1.TextMatrix(grid1.Rows - 1, 4) = locTable!TYPE_dESCA & ""
                grid1.TextMatrix(grid1.Rows - 1, 5) = locTable!Name & ""
                grid1.TextMatrix(grid1.Rows - 1, 6) = locTable!Phone & ""
                grid1.TextMatrix(grid1.Rows - 1, 7) = locTable!TOTAL & ""
                grid1.TextMatrix(grid1.Rows - 1, 8) = locTable!doc_no & ""
                grid1.TextMatrix(grid1.Rows - 1, 9) = locTable!Date & ""
                grid1.TextMatrix(grid1.Rows - 1, 10) = date_Pick
                grid1.TextMatrix(grid1.Rows - 1, 11) = cod
                grid1.TextMatrix(grid1.Rows - 1, 12) = locTable!Payment_Method & ""
                grid1.TextMatrix(grid1.Rows - 1, 13) = LCase(sType)
                grid1.TextMatrix(grid1.Rows - 1, 14) = locTable!ID
            Else
                grid2.AddItem ""
                grid2.TextMatrix(grid2.Rows - 1, 0) = grid2.Rows - 1
                grid2.TextMatrix(grid2.Rows - 1, 1) = locTable!ship_no & ""
                grid2.TextMatrix(grid2.Rows - 1, 2) = locTable!ORDER_NO & ""
                grid2.TextMatrix(grid2.Rows - 1, 3) = locTable!order_no_sup & ""
                grid2.TextMatrix(grid2.Rows - 1, 4) = locTable!TYPE_dESCA & ""
                grid2.TextMatrix(grid2.Rows - 1, 5) = locTable!Name & ""
                grid2.TextMatrix(grid2.Rows - 1, 6) = locTable!Phone & ""
                grid2.TextMatrix(grid2.Rows - 1, 7) = locTable!TOTAL & ""
                grid2.TextMatrix(grid2.Rows - 1, 8) = locTable!doc_no & ""
                grid2.TextMatrix(grid2.Rows - 1, 9) = locTable!Date & ""
                grid2.TextMatrix(grid2.Rows - 1, 10) = date_Pick
                grid2.TextMatrix(grid2.Rows - 1, 11) = cod
                grid2.TextMatrix(grid2.Rows - 1, 12) = locTable!Payment_Method & ""
                grid2.TextMatrix(grid2.Rows - 1, 13) = LCase(sType)
                grid2.TextMatrix(grid2.Rows - 1, 14) = locTable!ID
            End If
        Else
            grid3.AddItem ""
            grid3.TextMatrix(grid3.Rows - 1, 0) = grid3.Rows - 1
            grid3.TextMatrix(grid3.Rows - 1, 1) = ship_no
            grid3.TextMatrix(grid3.Rows - 1, 2) = date_Pick
            grid3.TextMatrix(grid3.Rows - 1, 3) = ref
            grid3.TextMatrix(grid3.Rows - 1, 4) = LCase(sType)
            grid3.TextMatrix(grid3.Rows - 1, 5) = sName
            grid3.TextMatrix(grid3.Rows - 1, 6) = Phone
            grid3.TextMatrix(grid3.Rows - 1, 7) = cod
         End If
    End If
Next
cleanUp:
grid1.Redraw = flexRDBuffered
grid2.Redraw = flexRDBuffered
grid3.Redraw = flexRDBuffered
prog1.Visible = False
Me.Caption = sCaption
Set db = Nothing
End Function
Private Sub Fixgrd1()
With grid1
.RowHeight(0) = 800
.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "»Ê·Ì’… «·‘Õ‰"
.TextMatrix(0, 2) = "—ﬁ„ «·ÿ·»Ì…"
.TextMatrix(0, 3) = "„” ‰œ ›—⁄Ì"
.TextMatrix(0, 4) = "‰Ê⁄ «·ÿ·»Ì…"
.TextMatrix(0, 5) = "«·«”„"
.TextMatrix(0, 6) = "«· ·Ì›Ê‰"
.TextMatrix(0, 7) = "≈Ã„«·Ì «·ÿ·»Ì…"
.TextMatrix(0, 8) = "—ﬁ„ „” ‰œ «·‘Õ‰"
.TextMatrix(0, 9) = " «—ÌŒ „” ‰œ «·‘Õ‰"
.TextMatrix(0, 10) = " «—ÌŒ «·‘Õ‰"
.TextMatrix(0, 11) = "COD"
.TextMatrix(0, 12) = "‰Ê⁄ «·”œ«œ"
.TextMatrix(0, 13) = "«·‰Ê⁄"

Dim i As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ColAlignment(3) = flexAlignLeftCenter

If .Rows > 1 Then
    .Cell(flexcpFontSize, 1, 13, 1, .Rows - 1) = 10
End If

For i = 1 To grid1.Rows - 1
    .TextMatrix(i, 0) = i
Next

.ColWidth(0) = 800
.ColWidth(1) = 1800
.ColWidth(2) = 1200
.ColWidth(3) = 1000
.ColWidth(4) = 1800
.ColWidth(5) = 3000
.ColWidth(6) = 2000
.ColWidth(7) = 1100
.ColWidth(8) = 900
.ColWidth(9) = 1250
.ColWidth(10) = 1250
.ColWidth(11) = 1200
.ColWidth(12) = 1500
.ColWidth(13) = 1500

.ColHidden(3) = True
.ColHidden(.Cols - 1) = True

.ColDataType(7) = flexDTDouble
.ColDataType(9) = flexDTDate
.ColDataType(10) = flexDTDate
.ColDataType(11) = flexDTDouble

.ColFormat(9) = "yyyy/M/d"
.ColFormat(10) = "yyyy/M/d"
End With
End Sub
Private Sub Fixgrd2()
With grid2
.RowHeight(0) = 800
.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "»Ê·Ì’… «·‘Õ‰"
.TextMatrix(0, 2) = "—ﬁ„ «·ÿ·»Ì…"
.TextMatrix(0, 3) = "„” ‰œ ›—⁄Ì"
.TextMatrix(0, 4) = "‰Ê⁄ «·ÿ·»Ì…"
.TextMatrix(0, 5) = "«·«”„"
.TextMatrix(0, 6) = "«· ·Ì›Ê‰"
.TextMatrix(0, 7) = "≈Ã„«·Ì «·ÿ·»Ì…"
.TextMatrix(0, 8) = "—ﬁ„ „” ‰œ «·‘Õ‰"
.TextMatrix(0, 9) = " «—ÌŒ „” ‰œ «·‘Õ‰"
.TextMatrix(0, 10) = " «—ÌŒ «·‘Õ‰"
.TextMatrix(0, 11) = "COD"
.TextMatrix(0, 12) = "‰Ê⁄ «·”œ«œ"
.TextMatrix(0, 13) = "«·‰Ê⁄"

Dim i As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

For i = 1 To grid1.Rows - 1
    .TextMatrix(i, 0) = i
Next

If .Rows > 1 Then
    .Cell(flexcpFontSize, 1, 13, 1, .Rows - 1) = 10
End If

.ColWidth(0) = 800
.ColWidth(1) = 1800
.ColWidth(2) = 1200
.ColWidth(3) = 1000
.ColWidth(4) = 1800
.ColWidth(5) = 3000
.ColWidth(6) = 2000
.ColWidth(7) = 1100
.ColWidth(8) = 900
.ColWidth(9) = 1250
.ColWidth(10) = 1250
.ColWidth(11) = 1200
.ColWidth(12) = 1500
.ColWidth(13) = 1500

.ColHidden(3) = True
.ColHidden(.Cols - 1) = True

.ColDataType(7) = flexDTDouble
.ColDataType(9) = flexDTDate
.ColDataType(10) = flexDTDate
.ColDataType(11) = flexDTDouble

.ColFormat(9) = "yyyy/M/d"
.ColFormat(10) = "yyyy/M/d"
End With
End Sub
Private Function myload3(ByRef cSv As ChilkatCsv) As Boolean
Dim i As Long
Dim sCaption As String
Dim ship_no As String
Dim date_Pick As String
Dim cod As String
Dim Tb As New ChilkatStringBuilder

sCaption = Me.Caption
grid3.Rows = 1

For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
        
    ship_no = addstring(Trim(cSv.GetCellByName(i, "Tracking Number")))
    
    date_Pick = Trim(cSv.GetCellByName(i, "Picked-Up Date"))
    date_Pick = Replace(date_Pick, ",", "")
    If IsDate(date_Pick) Then
        date_Pick = Format(date_Pick, "mm-dd-yyyy hh:nn")
        date_Pick = Format(date_Pick, "yyyy-mm-dd hh:nn")
    Else
        date_Pick = ""
    End If
    ref = addstring(cSv.GetCellByName(i, "Business Reference Number"))
    sType = addstring(cSv.GetCellByName(i, "Type"))
    sName = addstring(cSv.GetCellByName(i, "Consignee Name"))
    Phone = addstring(cSv.GetCellByName(i, "Consignee phone"))
    cod = Val(cSv.GetCellByName(i, "Cod Amount"))
    
    grid3.AddItem ""
    grid3.TextMatrix(grid3.Rows - 1, 0) = i
    grid3.TextMatrix(grid3.Rows - 1, 1) = sShip
    grid3.TextMatrix(grid3.Rows - 1, 2) = date_Pick
    grid3.TextMatrix(grid3.Rows - 1, 3) = sType
    grid3.TextMatrix(grid3.Rows - 1, 4) = sName
    grid3.TextMatrix(grid3.Rows - 1, 5) = Phone
    grid3.TextMatrix(grid3.Rows - 1, 5) = cod
Next



prog1.Visible = False
Me.Caption = sCaption
End Function
Private Sub Fixgrd3()
With grid3
.RowHeight(0) = 800
.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "»Ê·Ì’… «·‘Õ‰"
.TextMatrix(0, 2) = " «—ÌŒ «·‘Õ‰"
.TextMatrix(0, 3) = "Reference Number"
.TextMatrix(0, 4) = "Type"
.TextMatrix(0, 5) = "«·«”„"
.TextMatrix(0, 6) = "«· ·Ì›Ê‰"
.TextMatrix(0, 7) = "COD"

Dim i As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
Next

.ColWidth(0) = 800
.ColWidth(1) = 1800
.ColWidth(2) = 1300
.ColWidth(3) = 4000
.ColWidth(4) = 3500
.ColWidth(5) = 3500
.ColWidth(6) = 3000
.ColWidth(7) = 1200

.ColDataType(7) = flexDTDouble
.ColDataType(2) = flexDTDate

.ColFormat(2) = "yyyy/M/d"
End With
End Sub


