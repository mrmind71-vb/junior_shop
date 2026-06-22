VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form online_Show_Msg 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈ŸÂ«— «·—”«∆·"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   45
      TabIndex        =   3
      Top             =   3240
      Width           =   1545
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   600
         Left            =   45
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1058
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
         Picture         =   "online_show_msg.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   3210
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   7710
      Begin VB.TextBox xMsg 
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
         Height          =   2895
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   500
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   225
         Width           =   6810
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·»Ì«‰"
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
         Left            =   6975
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   645
      End
   End
End
Attribute VB_Name = "online_Show_Msg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sCaption As String
Public sMsg As String
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
If sCaption <> "" Then Me.Caption = sCaption
xMsg.text = sMsg
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set online_Show_Msg = Nothing
End Sub
Private Sub xMsg_GotFocus()
myGotFocus xMsg
End Sub
Private Sub xMsg_LostFocus()
myLostFocus xMsg
End Sub
