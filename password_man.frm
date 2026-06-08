VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Password_man 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ﬂÊœ «·„‰œÊ»"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   5295
      TabIndex        =   19
      Top             =   8310
      Width           =   5295
      Begin VB.Line Line2 
         BorderColor     =   &H00800000&
         BorderWidth     =   7
         X1              =   45
         X2              =   5535
         Y1              =   45
         Y2              =   45
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7485
      Left            =   0
      ScaleHeight     =   7485
      ScaleWidth      =   5295
      TabIndex        =   3
      Top             =   825
      Width           =   5295
      Begin VB.TextBox xPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F0EDED&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C4D3D&
         Height          =   510
         IMEMode         =   3  'DISABLE
         Left            =   180
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   405
         Width           =   3390
      End
      Begin Threed.SSCommand cmdNumber 
         Height          =   1140
         Index           =   2
         Left            =   1260
         TabIndex        =   4
         Top             =   1395
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":0000
         Caption         =   "2"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNumber 
         Height          =   1140
         Index           =   1
         Left            =   45
         TabIndex        =   5
         Top             =   1395
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":2D23
         Caption         =   "1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNumber 
         Height          =   1140
         Index           =   3
         Left            =   2475
         TabIndex        =   6
         Top             =   1395
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":5A46
         Caption         =   "3"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNumber 
         Height          =   1140
         Index           =   4
         Left            =   45
         TabIndex        =   7
         Top             =   2565
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":8769
         Caption         =   "4"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNumber 
         Height          =   1140
         Index           =   5
         Left            =   1260
         TabIndex        =   8
         Top             =   2565
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":B48C
         Caption         =   "5"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNumber 
         Height          =   1140
         Index           =   6
         Left            =   2475
         TabIndex        =   9
         Top             =   2565
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":E1AF
         Caption         =   "6"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNumber 
         Height          =   1140
         Index           =   7
         Left            =   45
         TabIndex        =   10
         Top             =   3780
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":10ED2
         Caption         =   "7"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNumber 
         Height          =   1140
         Index           =   8
         Left            =   1260
         TabIndex        =   11
         Top             =   3780
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":13BF5
         Caption         =   "8"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNumber 
         Height          =   1140
         Index           =   9
         Left            =   2475
         TabIndex        =   12
         Top             =   3780
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":16918
         Caption         =   "9"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdApply 
         Default         =   -1  'True
         Height          =   5955
         Left            =   3735
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1395
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   10504
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":1963B
         Caption         =   "Enter"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNumber 
         Height          =   1140
         Index           =   0
         Left            =   45
         TabIndex        =   14
         Top             =   4995
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":1DC49
         Caption         =   "0"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   1185
         Left            =   3735
         TabIndex        =   15
         Top             =   90
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   2090
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":2112E
         Caption         =   "Esc"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDel 
         Height          =   1140
         Left            =   45
         TabIndex        =   16
         Top             =   6210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":23E51
         Caption         =   "Del"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNumber 
         Height          =   1140
         Index           =   10
         Left            =   2475
         TabIndex        =   17
         Top             =   4995
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":26B74
         Caption         =   "."
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdBack 
         Height          =   1185
         Left            =   1260
         TabIndex        =   18
         Top             =   6165
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   2090
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   529464
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "password_man.frx":29897
         Caption         =   "Back Space"
         ButtonStyle     =   3
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         FillColor       =   &H00F0EDED&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   90
         Top             =   135
         Width           =   3570
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   5295
      TabIndex        =   1
      Top             =   0
      Width           =   5295
      Begin VB.Label xDate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   135
         TabIndex        =   2
         Top             =   135
         Width           =   5115
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   7
         X1              =   0
         X2              =   5490
         Y1              =   720
         Y2              =   720
      End
   End
End
Attribute VB_Name = "Password_man"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Form
Dim nTimes As Integer, nTime, userTable As Recordset
Private Sub CmdApply_Click()
If Trim(xPass.text) <> "" Then
    myForm.myProc
End If
End Sub
Private Sub cmdBack_Click()
xPass.SetFocus
SendKeys "{backspace}"
End Sub
Private Sub CmdDel_Click()
xPass.SetFocus
xPass.SelStart = 0
xPass.SelLength = Len(xPass.text)
SendKeys "{delete}"
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdNumber_Click(Index As Integer)
xPass.SetFocus
If cmdNumber(Index).Caption = "" Then
    xPass.text = ""
Else
    SendKeys IIf(bCapsMode, UCase(cmdNumber(Index).Caption), LCase(cmdNumber(Index).Caption))
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Then KeyAscii = 0
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Password_man = Nothing
End Sub

Private Sub xPass_KeyPress(KeyAscii As Integer)
    ' Allow numbers (0-9), Backspace (8), and Enter (13)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
    KeyAscii = 0 ' Cancel the keystroke
    'Beep         ' Optional: Make a sound on invalid key
End If
End Sub
