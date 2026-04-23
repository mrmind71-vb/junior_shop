VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form rpCharge_br 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4845
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   5820
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   600
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      Picture         =   "rpcharge_br.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   4050
      UseMaskColor    =   -1  'True
      Width           =   5550
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   135
      Width           =   5550
      Begin Threed.SSCommand cmdgo 
         Height          =   465
         Index           =   2
         Left            =   135
         TabIndex        =   1
         Top             =   765
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   820
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ›’Ì·Ì „’—Ê› ›—Ê⁄ Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   465
         Index           =   3
         Left            =   135
         TabIndex        =   2
         Top             =   1350
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   820
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "„ «»⁄… ÌÊ„Ì… ··„’«—Ì› ›—Ê⁄  ÊﬂÌ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   465
         Index           =   1
         Left            =   135
         TabIndex        =   0
         Top             =   180
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   820
         _Version        =   196610
         CaptionStyle    =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ì „’«—Ì› ›—Ê⁄  ÊﬂÌ·«  Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   465
         Index           =   4
         Left            =   135
         TabIndex        =   3
         Top             =   1935
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   820
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ï „’«—Ì› Õ”» «· ”ÃÌ·"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   465
         Index           =   5
         Left            =   135
         TabIndex        =   4
         Top             =   2520
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   820
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ï „Êﬁ› Õ”«»  ÊﬂÌ· "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   465
         Index           =   6
         Left            =   135
         TabIndex        =   5
         Top             =   3105
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   820
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ï „Êﬁ› Õ”«»  ÊﬂÌ·  - ·›—⁄ ·› —…"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "rpCharge_br"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOption As Integer
Private Sub cmdApply_Click()
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdGo_Click(Index As Integer)
publicFlag = Index
Select Case Index
Case 1
    rpcharge_br1.Show 1
Case 2
   rpcharge_BR2.Show 1
Case 3
    rpcharge_BR3.Show 1
Case 4
    rpcharge_br4.Show 1
Case 5
    rpcharge_br5.Show 1
Case 6
    rpcharge_br6.Show 1
End Select
End Sub
Private Sub cmdgo_MouseEnter(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdgo(Index).ForeColor = &HC00000
End Sub
Private Sub cmdgo_MouseExit(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdgo(Index).ForeColor = &H80000008
End Sub
Private Sub CmdOk_Click()
Unload Me
End Sub
