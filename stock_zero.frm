VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form stock_zero 
   BackColor       =   &H00FFFFFF&
   Caption         =   "«÷«›… «’‰«› ·Ì” ·Â« —’Ìœ"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1830
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame11 
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
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   990
      Width           =   2445
      Begin Threed.SSCommand cmdExit 
         Height          =   465
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   135
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
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
         Picture         =   "stock_zero.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   1170
         TabIndex        =   11
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   820
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
         Picture         =   "stock_zero.frx":2323
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "stock_zero.frx":4C48
      End
   End
   Begin VB.Frame Frame9 
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
      Height          =   960
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   8115
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   4500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xfact 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xmosm 
         Height          =   315
         Left            =   4500
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   540
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   315
         Left            =   90
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   540
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„’‰⁄"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   525
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄…"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   615
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
   Begin MSAdodcLib.Adodc DATA2 
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
   Begin MSAdodcLib.Adodc DATA3 
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
   Begin MSAdodcLib.Adodc DATA4 
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   12
      Top             =   1725
      Visible         =   0   'False
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "stock_zero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As ADODB.Connection
Public myform As Stock_appFrm
Public pZone As String
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not myform.myValid Then Exit Sub
If Not myform.myreplace(0) Then Exit Sub

Me.MousePointer = vbHourglass
Dim sError As String
Dim nRecords As Long

nRecords = AddZero(sError)
Me.MousePointer = vbNormal

If sError <> "" Then
    MsgBox sError
    Exit Sub
End If

If nRecords > 0 Then
    Inform "  „ «÷«›… " & nRecords & " ”Ã·"
    myform.myProcZero
End If
End Sub

Private Sub Form_Load()
Set data1.Recordset = mycmd("SELECT * FROM FILE1_10sc ORDER BY code ", con)
Set xSection.RowSource = data1
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

Set DATA2.Recordset = mycmd("SELECT * FROM fact ORDER BY code ", con)
Set xfact.RowSource = DATA2
xfact.ListField = "Desca"
xfact.BoundColumn = "Code"

Set DATA3.Recordset = mycmd("SELECT * FROM mosm ORDER BY date desc ", con)
Set xMosm.RowSource = DATA3
xMosm.ListField = "Desca"
xMosm.BoundColumn = "mosm"

Set data4.Recordset = mycmd("SELECT * from file1_50 ORDER BY desca ", con)
Set xgroup.RowSource = data4
xgroup.ListField = "Desca"
xgroup.BoundColumn = "Code"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set stockZero = Nothing
End Sub
Private Sub xSection_GotFocus()
myGotFocus xSection
End Sub
Private Sub xSection_LostFocus()
myLostFocus xSection
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
Private Sub xFact_GotFocus()
myGotFocus xfact
End Sub
Private Sub xFACT_LostFocus()
myLostFocus xfact
If Not xfact.MatchedWithList Then xfact.BoundText = ""
End Sub
Private Sub xMosm_GotFocus()
myGotFocus xMosm
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xMosm
If Not xMosm.MatchedWithList Then xMosm.BoundText = ""
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xgroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xgroup
If Not xgroup.MatchedWithList Then xgroup.BoundText = ""
End Sub
Private Function AddZero(pError As String) As Long
Dim loctable As New ADODB.Recordset
Dim cString As String

Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DOC_NO", myform.xdoc_no.text)
aPrm = AddFlag(aPrm, "MOSM", TurnValue(xMosm.BoundText))
aPrm = AddFlag(aPrm, "FACT", TurnValue(xfact.BoundText))
aPrm = AddFlag(aPrm, "SECTION", TurnValue(xSection.BoundText))
aPrm = AddFlag(aPrm, "GROUP", TurnValue(xgroup.BoundText))

Set loctable = cmd("dbo.sp_stock_zero_zone", con, adStoredProc, aPrm, 1000).Execute
If loctable.EOF And loctable.BOF Then
    MsgBox "·« ÌÊÃœ «’‰«› ··≈÷«›…"
    Exit Function
End If

Dim nRecordCount As Long, nRecord As Long
Dim sCaption As String
Dim nFound As Integer

nRecordCount = loctable.RecordCount

sCaption = Me.Caption
prog1.Visible = True
prog1.Value = 0
con.BeginTrans
'On Error GoTo myError

Dim sb As New ChilkatStringBuilder
Dim i As Long
Dim nAdded As Long
Do Until loctable.EOF
    Me.Caption = sCaption & " - " & "”Ã· " & loctable.AbsolutePosition & " „‰ " & loctable.RecordCount
    prog1.Value = Round(loctable.AbsolutePosition / loctable.RecordCount, 2) * 100
    
    i = i + 1
    sb.Append "(" & _
                addstring(myform.xdoc_no.text) & "," & _
                addstring(loctable!Item) & "," & _
                "0" & "," & _
                addstring(myform.xZone.Caption) & "," & _
                "1" & _
               "),"
    
    If i = 1000 Or loctable.AbsolutePosition = loctable.RecordCount Then
        sb.Shorten 1
        con.Execute "INSERT INTO FILE0_100(" & _
                    "doc_no," & _
                    " item," & _
                    " RealBal," & _
                    " [ZONE]," & _
                    " [AUTO])" & _
                     "VALUES" & _
                     sb.GetAsString, nAdded
        sb.Clear
        nAffect = nAffect + nAdded
        i = 0
    End If
    
    loctable.MoveNext
Loop
AddZero = nAffect
con.CommitTrans
Finally:
prog1.Visible = False
Exit Function
myError:
pError = Err.Description
AddZero = -1
Err.Clear
con.RollbackTrans
End Function
