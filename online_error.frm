VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form online_error 
   BackColor       =   &H00FFFFFF&
   Caption         =   "„‘ﬂ·«  «·«Ê‰·«Ì‰"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   45
      TabIndex        =   3
      Top             =   1485
      Width           =   7710
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   555
         Left            =   45
         TabIndex        =   4
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
         Picture         =   "online_error.frx":0000
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
         TabIndex        =   5
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
         Picture         =   "online_error.frx":2323
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_error.frx":4C48
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   7755
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
         Height          =   1185
         Left            =   90
         MaxLength       =   200
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Tag             =   "22"
         Top             =   180
         Width           =   6135
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
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1275
      End
   End
End
Attribute VB_Name = "online_error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Form
Public sCaption As String
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Trim(xMsg.text) <> "" Then
    myForm.sMsg = xMsg.text
    Unload Me
End If
End Sub

Private Sub Form_Load()
If sCaption <> "" Then Me.Caption = sCaption
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set online_error = Nothing
End Sub

Private Sub xMsg_GotFocus()
myGotFocus xMsg
End Sub

Private Sub xMsg_LostFocus()
myLostFocus xMsg
End Sub
