VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form stock_zonefrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   " ⁄œÌ· «·“Ê‰"
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
      TabIndex        =   5
      Top             =   1215
      Width           =   2445
      Begin Threed.SSCommand cmdExit 
         Height          =   465
         Left            =   0
         TabIndex        =   2
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
         Picture         =   "stock_zone.frx":0000
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
         TabIndex        =   1
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
         Picture         =   "stock_zone.frx":2323
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "stock_zone.frx":4C48
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
      TabIndex        =   3
      Top             =   45
      Width           =   6090
      Begin VB.TextBox xZone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   225
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   630
         Width           =   4065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œÌ· ≈·Ì"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   690
      End
      Begin VB.Label xZone_org 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   4065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·“Ê‰"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   765
      End
   End
End
Attribute VB_Name = "stock_zonefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As ADODB.Connection
Public myform As Form
Public bMerge As Boolean
Public pDoc_no As String
Public pZone As String
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Trim(xZone.text) = Trim(xZone_org.Caption) Then
    MsgBox "‰›” «·“Ê‰ «·„—«œ  ⁄œÌ·Â"
    Exit Sub
End If

Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "SELECT TOP 1 ID " & _
          " FROM FILE0_100 " & _
          " WHERE ZONE = " & MyParn(xZone.text) & _
          " AND DOC_NO = " & MyParn(pDoc_no)
          
Set loctable = cmd(cString, con).Execute

If (Not loctable.EOF) Then
    MsgBox "—ﬁ„ «·“Ê‰ „ÊÃÊœ"
    Unload Me
End If
myform.myProcZoneEdit
End Sub
Private Sub Form_Load()
xZone_org.Caption = pZone
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set stock_zonefrm = Nothing
End Sub

Private Sub xZone_Change()
cmdSave.Enabled = Trim(xZone.text) <> ""
End Sub
Private Sub xZone_GotFocus()
myGotFocus xZone
End Sub
Private Sub xZone_LostFocus()
myLostFocus xZone
End Sub


