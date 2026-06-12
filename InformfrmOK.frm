VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form InformfrmOK 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5835
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand cmdOk 
      Default         =   -1  'True
      Height          =   1095
      Left            =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1931
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
      Picture         =   "InformfrmOK.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin VB.Label lbl_inform 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   1080
      Width           =   5685
   End
   Begin VB.Label lbl_inform 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1140
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   270
      Width           =   5685
   End
End
Attribute VB_Name = "InformfrmOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCaption1 As String
Public sCaption2 As String

Private Sub cmdOk_Click()
    Unload Me
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
lbl_inform(0).Caption = sCaption1
lbl_inform(1).Caption = sCaption2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set InformfrmOK = Nothing
End Sub

