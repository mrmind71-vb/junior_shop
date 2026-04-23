VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form stock_zone_addfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "«÷«›… “Ê‰"
   ClientHeight    =   1485
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
   ScaleHeight     =   1485
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
      Top             =   765
      Width           =   2445
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
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
         Picture         =   "stock_zone_add.frx":0000
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
         Picture         =   "stock_zone_add.frx":2323
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "stock_zone_add.frx":4C48
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
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   45
      Width           =   6090
      Begin VB.TextBox xZone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   135
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   4065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·“Ê‰"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   765
      End
   End
End
Attribute VB_Name = "stock_zone_addfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As ADODB.Connection
Public myform As Form
Public pDoc_no As String
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "SELECT TOP 1 ID " & _
          " FROM FILE0_100 " & _
          " WHERE ZONE = " & MyParn(xZone.text) & _
          " AND DOC_NO = " & MyParn(pDoc_no)

Set loctable = cmd(cString, con).Execute
If Not loctable.EOF Then
    MsgBox "—ﬁ„ «·“Ê‰ „ÊÃÊœ „‰ ﬁ»·"
    Exit Sub
End If

myform.myProcZoneAdd
End Sub
Private Sub Form_Load()
Dim loctable As New ADODB.Recordset
Set loctable = cmd("select max(zone) as zone from file0_100 where doc_no = " & MyParn(pDoc_no), con).Execute
If IsNull(loctable!ZONE) Then
    xZone.text = "Z1"
Else
    xZone.text = IncrementString(loctable!ZONE)
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set stock_zone_addfrm = Nothing
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
Private Function IncrementString(ByVal inputString As String) As String
    Dim i As Long
    Dim numericPart As String
    Dim alphaPart As String
    
    ' Handle empty string case
    If Len(inputString) = 0 Then
        IncrementString = "1"
        Exit Function
    End If
    
    ' Find the start of the numeric part at the end of the string
    For i = Len(inputString) To 1 Step -1
        If Not IsNumeric(Mid(inputString, i, 1)) Then
            Exit For
        End If
    Next i
    
    ' Split string into alphabet and numeric parts
    alphaPart = Left(inputString, i)
    numericPart = Right(inputString, Len(inputString) - i)
    
    If numericPart = "" Then
        ' Scenario 1: Ends with a character (e.g., "z" -> "z2")
        IncrementString = alphaPart & "2"
    Else
        ' Scenario 2: Ends with a number (e.g., "z9" -> "z10")
        ' Val() converts string to number, then we add 1 and concatenate back
        IncrementString = alphaPart & (Val(numericPart) + 1)
    End If
End Function

