VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Phonesfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   " ·Ì›Ê‰«  «·⁄„·«¡"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1935
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   90
      TabIndex        =   6
      Top             =   1080
      Width           =   7710
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   135
         Width           =   1410
         _ExtentX        =   2487
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
         Picture         =   "Phones.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSave 
         Default         =   -1  'True
         Height          =   555
         Left            =   1485
         TabIndex        =   8
         Top             =   135
         Width           =   1410
         _ExtentX        =   2487
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
         Picture         =   "Phones.frx":2323
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Phones.frx":4C48
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   7755
      Begin VB.TextBox xPhone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         MaxLength       =   11
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "22"
         Top             =   180
         Width           =   5190
      End
      Begin VB.TextBox xName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "22"
         Top             =   585
         Width           =   5190
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„  ·Ì›Ê‰ «·⁄„Ì·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«”„ «·⁄„Ì·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   630
         Width           =   1365
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VSFlex7Ctl.VSFlexGrid grdPhone 
      Height          =   3795
      Left            =   45
      TabIndex        =   1
      Top             =   630
      Visible         =   0   'False
      Width           =   7755
      _cx             =   13679
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
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
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
      TabBehavior     =   0
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
Attribute VB_Name = "Phonesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As ADODB.Connection
Public myform As Form
Dim bIg As Boolean
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Trim(XPHONE.text) <> "" Then
    If Not IsValidMobile(XPHONE.text) Then
        MsgBox "—ﬁ„ „Õ„Ê· €Ì— ’«·Õ"
        Exit Sub
    End If
    
    myform.XPHONE.Caption = XPHONE.text
    If Trim(xName.text) <> "" Then
        myform.XPHONE.Tag = xName.text
    Else
        xName.text = fnPhoneName(XPHONE.text, con)
    End If
    Unload Me
End If
End Sub
Private Sub Form_Load()
Set grdPhone.DataSource = data1
If myform.XPHONE.Caption <> "" Then
    bIg = True
    XPHONE.text = myform.XPHONE.Caption
    If myform.XPHONE.Tag <> "" Then
        xName.text = myform.XPHONE.Tag
    Else
        xName.text = fnPhoneName(XPHONE.text, con)
    End If
    bIg = False
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Phonesfrm = Nothing
End Sub

Private Sub xPhone_Change()
If bIg Then Exit Sub
myLoadList , XPHONE.text
End Sub
Private Sub myLoadList(Optional pName As String = "", Optional pPhone As String = "")
Dim cString As String

If Not grdPhone.Visible Then
    grdPhone.Visible = True
    Me.Height = Me.Height + 3000
End If

cString = "Select top 100 Phone," & _
          "desca," & _
          "format(f_date,'yyyy/M/d')" & _
          " FROM SUBCUST"
          
If Trim(pName) <> "" Then
    cString = cString & " WHERE DESCA LIKE " & MyParn(pName & "%")
ElseIf Trim(pPhone) <> "" Then
    cString = cString & " WHERE PHONE LIKE " & MyParn(pPhone & "%")
Else
    cString = cString & " WHERE (NOT (DESCA IS NULL AND PHONE IS NULL))"
End If
cString = cString & " ORDER BY " & IIf(Trim(pName) <> "", "[desca]", "Phone")
Set data1.Recordset = cmd(cString, con).Execute
FixGrdPhone
If grdPhone.Rows > 1 And grdPhone.Row < 1 Then
    grdPhone.Row = 1
ElseIf grdPhone.Rows = 1 Then
    grdPhone.Visible = False
    Me.Height = Me.Height - 3000
End If
End Sub
Private Sub FixGrdPhone()
grdPhone.RowHidden(0) = True
grdPhone.ColWidth(0) = 2000
grdPhone.ColWidth(1) = 2500
grdPhone.ColWidth(2) = 1300

If grdPhone.Rows > 1 Then
    grdPhone.Cell(flexcpBackColor, 1, 2, grdPhone.Rows - 1, 2) = &H8000000F
    GrdHi grdPhone, 2, -1, grdPhone.Row
End If

For i = 0 To grdPhone.Cols - 1
    grdPhone.ColAlignment(i) = flexAlignRightCenter
Next
End Sub
Private Sub grdPhone_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    grdPhone_DblClick
ElseIf KeyCode = 27 Then
    grdPhone.Visible = False
    Me.Height = Me.Height - 3000
End If
End Sub
Private Sub grdPhone_DblClick()
bIg = True
XPHONE.text = grdPhone.TextMatrix(grdPhone.Row, 0)
xName.text = grdPhone.TextMatrix(grdPhone.Row, 1)
'xE_mail.text = grdPhone.TextMatrix(grdPhone.Row, 2)
grdPhone.Visible = False
Me.Height = Me.Height - 3000
bIg = False
xName.SetFocus
End Sub
Public Sub GrdHi(pGrd As Object, pCol As Long, pRow1 As Long, pRow2 As Long, Optional pColor As Long = &HC0FFFF)
    If pCol > 0 Then
        If pRow1 <> -1 Then pGrd.Cell(flexcpBackColor, pRow1, 0, pRow1, pCol - 1) = pGrd.BackColor
        If pRow2 <> -1 Then pGrd.Cell(flexcpBackColor, pRow2, 0, pRow2, pCol - 1) = pColor
    End If
    
    If pCol < pGrd.Cols - 1 Then
        If pRow1 <> -1 Then pGrd.Cell(flexcpBackColor, pRow1, pCol + 1, pRow1, pGrd.Cols - 1) = pGrd.BackColor
        If pRow2 <> -1 Then pGrd.Cell(flexcpBackColor, pRow2, pCol + 1, pRow2, pGrd.Cols - 1) = pColor
    End If
End Sub
Private Sub xPhone_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If grdPhone.Visible Then
        grdPhone.Visible = False
        Me.Height = Me.Height - 3000
    End If
End If
End Sub

Private Sub XPHONE_LostFocus()
myLostFocus XPHONE
End Sub
