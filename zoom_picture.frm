VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form BuildZoomForm 
   Caption         =   "Build Zoom.txt for Zoom In On Map"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "BuildZoomForm"
   ScaleHeight     =   5955
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBase 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5460
      Left            =   5625
      ScaleHeight     =   5400
      ScaleWidth      =   7200
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   7260
   End
   Begin VB.PictureBox picZoom 
      AutoRedraw      =   -1  'True
      Height          =   5460
      Left            =   5400
      ScaleHeight     =   5400
      ScaleWidth      =   9495
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   900
      Width           =   9555
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Point"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   1620
      Width           =   1035
   End
   Begin VB.Timer tmrResizeHelp 
      Enabled         =   0   'False
      Left            =   1980
      Top             =   1620
   End
   Begin VB.CommandButton cmdZoom100 
      Caption         =   "Zoom 100"
      Height          =   495
      Left            =   2700
      TabIndex        =   4
      Top             =   4140
      Width           =   1035
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2700
      TabIndex        =   6
      Top             =   4740
      Width           =   1035
   End
   Begin VB.ListBox lstNames 
      Height          =   1860
      IntegralHeight  =   0   'False
      Left            =   990
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   3675
   End
   Begin VB.CommandButton cmdSaveToDisk 
      Caption         =   "Save Changes To Disk"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   5340
      Width           =   2415
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   4740
      Width           =   1035
   End
   Begin VB.ComboBox cboZoomBy 
      Height          =   330
      ItemData        =   "zoom_picture.frx":0000
      Left            =   1320
      List            =   "zoom_picture.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtY 
      Height          =   330
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   3780
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   1320
      TabIndex        =   0
      Top             =   2940
      Width           =   2415
   End
   Begin VB.TextBox txtX 
      Height          =   330
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   3360
      Width           =   1215
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6810
      Left            =   1980
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3150
      Width           =   5010
      _cx             =   8837
      _cy             =   12012
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
      BackColorBkg    =   -2147483636
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
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
      Editable        =   2
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
   Begin VB.Label lblPicZoom 
      Alignment       =   2  'Center
      Caption         =   "Drag Image To Center Zoom Point"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Current Zoom Points"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3555
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Zoom By %"
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   4260
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Y"
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   3840
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   3000
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "X"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   3420
      Width           =   1155
   End
End
Attribute VB_Name = "BuildZoomForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const UseWMF = True

#If UseWMF Then
Private Const IMAGE_FILE As String = "Island.wmf" 'Better scaling capabilities.
#Else
Private Const IMAGE_FILE As String = "Island.gif"
#End If

Private Const ZOOM_FILE As String = "Zoom.txt"

Private Const FORM_MINWIDTH = 9000
Private Const FORM_MINHEIGHT = 6495

Private Declare Function GetKeyState Lib "user32" ( _
    ByVal nVirtKey As KeyCodeConstants) As Integer

Private ZoomPoints As Collection
Private CurrentZoomFactor As Single
Private CurrentBaseX As Single
Private CurrentBaseY As Single
Private ClipWidth As Single
Private ClipHeight As Single
Private DragX As Single
Private DragY As Single
Private SizeRatioX As Single
Private SizeRatioY As Single
Private AspectRatioPicZoom As Single
Private PaddingRightPicZoom As Single
Private PaddingBottomPicZoom As Single
Private Zoom100PctCboZoomIndex As Integer

'Event handler modifiers:
Private ClearingLstNamesSelection As Boolean
Private LoadingCboZoomBy As Boolean
Private NoFocusCboZoomBy As Boolean

Private Sub ResetUI(Optional ByVal FormLoad As Boolean)
    Zoom100Percent
    
    ClearingLstNamesSelection = True
    lstNames.ListIndex = -1
    ClearingLstNamesSelection = False
    
    txtName.text = ""
    txtX.text = CStr(Int(CurrentBaseX))
    txtY.text = CStr(Int(CurrentBaseY))
    
    If Not FormLoad Then txtName.SetFocus
End Sub

Private Sub SetClip(ByVal ZoomBy As Integer)
    CurrentZoomFactor = CSng(ZoomBy) / 100#
    With picBase
        ClipWidth = .ScaleWidth / CurrentZoomFactor
        ClipHeight = .ScaleHeight / CurrentZoomFactor
    End With
    If ClipWidth / ClipHeight < AspectRatioPicZoom Then
        ClipHeight = ClipWidth / AspectRatioPicZoom
    Else
        ClipWidth = ClipHeight * AspectRatioPicZoom
    End If
End Sub

Private Sub Zoom100Percent()
    SetClip 100
    CurrentBaseX = picBase.ScaleWidth / 2#
    CurrentBaseY = picBase.ScaleHeight / 2#
    ZoomPaint
    txtX.text = CStr(Int(CurrentBaseX))
    txtY.text = CStr(Int(CurrentBaseY))
    
    cboZoomBy.ListIndex = Zoom100PctCboZoomIndex
End Sub

Private Sub ZoomPaint()
    With picZoom
        Set .Picture = Nothing
        .PaintPicture picBase.Picture, _
                      0, 0, _
                      .ScaleWidth, .ScaleHeight, _
                      CurrentBaseX - (ClipWidth / 2#), CurrentBaseY - (ClipHeight / 2#), _
                      ClipWidth, ClipHeight
    End With
End Sub

Private Function ZoomPointExists(ByVal Name As String) As Boolean
    Dim ZP As ZoomPoint
    
    'Though stored Name value maintains case, key values are
    'case-insensitive so we must use a case-insensitive test.
    Name = UCase$(Name)
    For Each ZP In ZoomPoints
        If Name = UCase$(ZP.Name) Then
            ZoomPointExists = True
            Exit Function
        End If
    Next
End Function

Private Sub cboZoomBy_Click()
    If LoadingCboZoomBy Then Exit Sub
    
    SetClip CInt(cboZoomBy.text)
    ZoomPaint

    If NoFocusCboZoomBy Then Exit Sub
    
    txtName.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim Name As String
    Dim ZP As ZoomPoint
    Dim Index As Integer
    
    Name = UCase$(Trim$(txtName.text))
    Set ZP = ZoomPoints.Item(Name)
    ZoomPoints.Remove ZP.Name
    With lstNames
        For Index = 0 To .ListCount - 1
            If Name = UCase$(.List(Index)) Then
                .RemoveItem Index
                Exit For
            End If
        Next
    End With
    
    cmdSaveToDisk.Enabled = True
    
    ResetUI
End Sub

Private Sub cmdNew_Click()
    ResetUI
End Sub

Private Sub cmdSaveToDisk_Click()
    Dim F As Integer
    Dim ZP As ZoomPoint
    
    F = FreeFile(0)
    Open ZOOM_FILE For Output As #F
    For Each ZP In ZoomPoints
        With ZP
            Write #F, .Name, .X, .Y, .ZoomFactor
        End With
    Next
    Close #F
    cmdSaveToDisk.Enabled = False
    
    txtName.SetFocus
End Sub

Private Sub cmdUpdate_Click()
    Dim ZP As ZoomPoint
    
    cmdSaveToDisk.Enabled = True
    
    With lstNames
        If .ListIndex > -1 Then
            Set ZP = ZoomPoints.Item(.List(.ListIndex))
            ZoomPoints.Remove ZP.Name
            .RemoveItem .ListIndex
        Else
            Set ZP = New ZoomPoint
        End If
    End With
    
    With ZP
        .Name = Trim$(txtName.text)
        .X = CSng(txtX.text)
        .Y = CSng(txtY.text)
        .ZoomFactor = CSng(cboZoomBy.text) / 100#
        ZoomPoints.Add ZP, .Name
        lstNames.AddItem .Name
    End With
    
    ResetUI
End Sub

Private Sub cmdZoom100_Click()
    Zoom100Percent
    
    txtName.SetFocus
End Sub

Private Sub Form_Load()
    Dim F As Integer
    Dim ZP As ZoomPoint
    Dim Name As String
    Dim X As Single
    Dim Y As Single
    Dim ZoomFactor As Single
    Dim I As Long
    
    With tmrResizeHelp
        .Enabled = False
        .Interval = 100
    End With
    
    Set ZoomPoints = New Collection
    On Error Resume Next
    GetAttr ZOOM_FILE
    If Err.Number = 0 Then
        On Error GoTo 0
        F = FreeFile(0)
        Open ZOOM_FILE For Input As #F
        Do Until EOF(F)
            Set ZP = New ZoomPoint
            With ZP
                Input #F, Name, X, Y, ZoomFactor
                .Name = Name
                .X = X
                .Y = Y
                .ZoomFactor = ZoomFactor
                ZoomPoints.Add ZP, .Name
                lstNames.AddItem .Name
            End With
        Loop
        Close #F
        lstNames.ListIndex = -1
    Else
        On Error GoTo 0
    End If
    
    LoadingCboZoomBy = True
    With cboZoomBy
        .Clear
        .AddItem "50"
        .AddItem "75"
        'Some fiddly code to ensure we always have 100% in the list in its
        'correct position even if our loop doesn't hit I = 100.  Also
        'record its index for use in Zoom100Percent.
        Zoom100PctCboZoomIndex = -1
        For I = 100 To 1000 Step 100
            If I > 100 And Zoom100PctCboZoomIndex < 0 Then
                .AddItem "100"
                Zoom100PctCboZoomIndex = .NewIndex
            End If
            .AddItem CStr(I)
            If I = 100 Then Zoom100PctCboZoomIndex = .NewIndex
        Next
        .ListIndex = Zoom100PctCboZoomIndex
    End With
    LoadingCboZoomBy = False
    
    Set picBase.Picture = LoadPicture("C:\WORK\PRGSQL\Junior\JENUR_SHOP\documents\250812-480-124-56.jpg")
    
    'Set picZoom's backdrop color, get aspect ratio and padding for resize event:
    With picZoom
        .BackColor = picBase.Point(0, 0)
        If .BackColor = picBase.BackColor Then .BackColor = &H808080
        
        PaddingRightPicZoom = ScaleWidth - .Left - .Width
        PaddingBottomPicZoom = ScaleHeight - .Top - .Height
        AspectRatioPicZoom = .ScaleWidth / .ScaleHeight
    End With
    
    ResetUI FormLoad:=True
End Sub

Private Sub Form_Resize()
    Dim PicZoomNewWidth As Single
    Dim PicZoomNewHeight As Single
    
    If WindowState <> vbMinimized Then
        'Min size control logic:
        If Width < FORM_MINWIDTH Or Height < FORM_MINHEIGHT Then
            Refresh
            tmrResizeHelp.Enabled = True
            Exit Sub
        End If
        With picZoom
            'Maintain aspect ratio of picZoom:
            PicZoomNewWidth = ScaleWidth - .Left - PaddingRightPicZoom
            PicZoomNewHeight = ScaleHeight - .Top - PaddingBottomPicZoom
            If PicZoomNewWidth / PicZoomNewHeight < AspectRatioPicZoom Then
                PicZoomNewHeight = PicZoomNewWidth / AspectRatioPicZoom
            Else
                PicZoomNewWidth = PicZoomNewHeight * AspectRatioPicZoom
            End If
            
            lblPicZoom.Width = PicZoomNewWidth
            .Move .Left, .Top, PicZoomNewWidth, PicZoomNewHeight
            
            'Since picBase.AutoSize = True the loaded picture causes it
            'to resize to fit.  We'll need this ratio to scale drag
            'operations:
            .Refresh
            SizeRatioX = picBase.ScaleWidth / .ScaleWidth
            SizeRatioY = picBase.ScaleHeight / .ScaleHeight
            ZoomPaint
        End With
    End If
End Sub

Private Sub lstNames_Click()
    Dim ZP As ZoomPoint
    
    If ClearingLstNamesSelection Then Exit Sub
    
    Set ZP = ZoomPoints.Item(lstNames.List(lstNames.ListIndex))
    With ZP
        txtName.text = .Name
        CurrentBaseX = .X
        CurrentBaseY = .Y
        txtX.text = CStr(Int(.X))
        txtY.text = CStr(Int(.Y))
        NoFocusCboZoomBy = True
        cboZoomBy.text = CStr(Int(.ZoomFactor * 100#))
        NoFocusCboZoomBy = False
        SetClip CInt(cboZoomBy.text)
    End With
    cmdDelete.Enabled = True
End Sub

Private Sub picZoom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragX = X
    DragY = Y
End Sub

Private Sub picZoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CurrentBaseX = CurrentBaseX _
                     + (DragX - X) * SizeRatioX / CurrentZoomFactor
        CurrentBaseY = CurrentBaseY _
                     + (DragY - Y) * SizeRatioY / CurrentZoomFactor
        txtX.text = CStr(Int(CurrentBaseX))
        txtY.text = CStr(Int(CurrentBaseY))
        DragX = X
        DragY = Y
        ZoomPaint
    End If
End Sub

Private Sub tmrResizeHelp_Timer()
    'Exit while left mouse button is still down:
    If GetKeyState(vbKeyLButton) < 0 Then Exit Sub
    
    tmrResizeHelp.Enabled = False
    
    'Exit if minimized:
    If WindowState = vbMinimized Then Exit Sub
    
    If Width < FORM_MINWIDTH Then Width = FORM_MINWIDTH
    If Height < FORM_MINHEIGHT Then Height = FORM_MINHEIGHT
End Sub

Private Sub txtName_Change()
    cmdUpdate.Enabled = Len(Trim$(txtName.text)) > 0
    cmdDelete.Enabled = ZoomPointExists(Trim$(txtName.text))
End Sub
