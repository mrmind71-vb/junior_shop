VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form einvoice_menufrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "«·›« Ê—… «·«·Ìﬂ —Ê‰Ì…"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2880
      Width           =   4650
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·«Ì’«· «·«·Ìﬂ —Ê‰Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   1770
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·›« Ê—… «·«·Ìﬂ —Ê‰Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "÷»ÿ «⁄œ«œ  «·›« Ê—… «·«·Ìﬂ —Ê‰Ì…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   4650
      Begin Threed.SSCommand cmdSetting 
         Height          =   555
         Left            =   45
         TabIndex        =   1
         Top             =   360
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   979
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "÷»ÿ «·«⁄œ«œ« "
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdCompany 
         Height          =   555
         Left            =   45
         TabIndex        =   2
         Top             =   945
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   979
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "÷»ÿ »Ì«‰«  «·‘—ﬂ…"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   3
         Top             =   2115
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   979
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Œ—ÊÃ"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdRegister 
         Height          =   555
         Left            =   45
         TabIndex        =   7
         Top             =   1530
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   979
         _Version        =   196610
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ›⁄Ì· ›—⁄ «·«Ì’«·"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "einvoice_menufrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompany_Click()
If optType(0).Value Then
    einvoice_companyfrm.Show 1
Else
    receipt_companyfrm.Show 1
End If
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdRegister_Click()
Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")

Dim sError As String
Dim sToken As String
sToken = getToken_rc(sError)

If sError <> "" Then
    MsgBox sError
ElseIf sToken = "" Then
    MsgBox "Empty token"
Else
    MsgBox "Token : " & sToken
End If
End Sub

Private Sub cmdSetting_Click()
If optType(0).Value Then
    einvoice_settingfrm.Show 1
Else
    receipt_settingfrm.Show 1
End If
End Sub
Private Sub Form_Load()
If sBranchCode <> "00" Then
    optType(0).Enabled = False
    optType(1).Value = True
End If
End Sub

Private Sub optType_Click(Index As Integer)
Frame1.Caption = "÷»ÿ «⁄œ«œ«  " & optType(Index).Caption
cmdRegister.Enabled = Index = 1
End Sub

