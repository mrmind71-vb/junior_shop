VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form subUserPasswordrfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ﬂ·„… «·”—"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6165
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
   ScaleHeight     =   1845
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   -45
      Width           =   5955
      Begin VB.TextBox xPassword 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   405
         Width           =   4020
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬂ·„… «·”—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   450
         Width           =   1185
      End
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1138
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
      Picture         =   "subUserPassword.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdOk 
      Default         =   -1  'True
      Height          =   645
      Left            =   1755
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1138
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
      Picture         =   "subUserPassword.frx":2323
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
End
Attribute VB_Name = "subUserPasswordrfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdOk_Click()
Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "SELECT CODE," & _
          "DESCA FROM SUB_USER" & _
          " WHERE ISRETURN = 1" & _
          " AND " & _
          "(" & _
          "(PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0)" & _
          " OR " & _
          "(PASSWORD2 = " & MyParn(cPassword) & " AND STOP2 = 0)" & _
          ")"
Set loctable = cmd(cString, con).Execute
If Not loctable.EOF Then
    myform.subUserCode = loctable!Code
    myform.subUserName = loctable!DESCA & ""
    Unload Me
Else
    MsgBox "ﬂ·„… ”— €Ì— ’ÕÌÕ… «Ê ·Ì”  ·œÌﬂ «·’·«ÕÌ…"
End If
End Sub
Private Sub Form_Load()
openCon con
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set subUserpasswordFrm = Nothing
End Sub
Private Sub xPassword_Change()
cmdOk.Enabled = Trim(xPassword.text) = ""
End Sub
