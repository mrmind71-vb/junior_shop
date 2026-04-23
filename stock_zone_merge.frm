VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form stock_zone_mergefrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "œ„Ã «·“Ê‰"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6375
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
   ScaleHeight     =   1920
   ScaleWidth      =   6375
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
      TabIndex        =   4
      Top             =   1215
      Width           =   2445
      Begin Threed.SSCommand cmdExit 
         Height          =   465
         Left            =   0
         TabIndex        =   1
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
         Picture         =   "stock_zone_merge.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSave 
         Default         =   -1  'True
         Height          =   465
         Left            =   1170
         TabIndex        =   0
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
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
         Picture         =   "stock_zone_merge.frx":2323
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "stock_zone_merge.frx":4C48
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
      Height          =   1140
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   6090
      Begin MSDataListLib.DataCombo xZone 
         Height          =   330
         Left            =   225
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   630
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "œ„Ã „⁄"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   540
      End
      Begin VB.Label xZone_org 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   4065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·“Ê‰"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   765
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
End
Attribute VB_Name = "stock_zone_mergefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As ADODB.Connection
Public myForm As Form
Public bMerge As Boolean
Public pDoc_no As String
Public pZone As String
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Trim(xZone.BoundText) = "" Then
    MsgBox "«·“Ê‰ €Ì— „”Ã·"
    Exit Sub
End If

Dim cString As String
cString = "SELECT TOP 1 ID " & _
          " FROM FILE0_100 " & _
          " WHERE ZONE = " & MyParn(xZone.BoundText) & _
          " AND DOC_NO = " & MyParn(pDoc_no)
Dim loctable As New ADODB.Recordset
Set loctable = cmd(cString, con).Execute
If (loctable.EOF) Then
    MsgBox "«·“Ê‰ «·„›—Ê÷ «·œ„Ã „⁄Â €Ì— „ÊÃÊœ"
    Unload Me
End If
myForm.myProcZoneMerge
End Sub
Private Sub Form_Load()
Set DATA1.Recordset = mycmd("SELECT ZONE FROM FILE0_100 WHERE DOC_NO = " & MyParn(pDoc_no) & " AND ZONE <> " & MyParn(pZone) & " GROUP BY ZONE ORDER BY ZONE", con)
Set xZone.RowSource = DATA1
xZone.ListField = "ZONE"
xZone.BoundColumn = "ZONE"

xZone_org.Caption = pZone
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set stock_zone_mergefrm = Nothing
End Sub

Private Sub xZone_Change()
'cmdSave.Enabled = xZone.MatchedWithList
End Sub

Private Sub xZone_Click(Area As Integer)
cmdSave.Enabled = xZone.MatchedWithList
End Sub

Private Sub xZone_GotFocus()
'myGotFocus xZone
End Sub
Private Sub xZone_LostFocus()
'myLostFocus xZone
'If Not xZone.MatchedWithList Then xZone.BoundText = ""
End Sub

