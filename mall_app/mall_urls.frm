VERSION 5.00
Begin VB.Form mall_settingfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "÷»ÿ «⁄œ«œ«   ÿ»Ìﬁ «·„Ê·"
   ClientHeight    =   2850
   ClientLeft      =   -15
   ClientTop       =   435
   ClientWidth     =   11685
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
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetToken 
      Caption         =   "Get Token"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7290
      MaskColor       =   &H00FFFFFF&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Õ›Ÿ"
      Top             =   2160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Frame Frame2 
      Height          =   2040
      Left            =   45
      TabIndex        =   7
      Top             =   45
      Width           =   11445
      Begin VB.TextBox xPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   4
         Top             =   1620
         Width           =   8430
      End
      Begin VB.TextBox xMail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   3
         Top             =   1260
         Width           =   8430
      End
      Begin VB.TextBox xInvoiceUrl 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   2
         Top             =   900
         Width           =   8430
      End
      Begin VB.TextBox xBaseUrl 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   0
         Top             =   180
         Width           =   8430
      End
      Begin VB.TextBox xAuthUrl 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   1
         Top             =   540
         Width           =   8430
      End
      Begin VB.Label Label3 
         Caption         =   "Invoices URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Mail Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Auth URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   585
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         TabIndex        =   9
         Top             =   1620
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Base URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdExit 
      Height          =   555
      Left            =   10170
      MaskColor       =   &H00FFFFFF&
      Picture         =   "mall_urls.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8730
      MaskColor       =   &H00FFFFFF&
      Picture         =   "mall_urls.frx":4AB6
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Õ›Ÿ"
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1410
   End
End
Attribute VB_Name = "mall_settingfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub cmdGetToken_Click()
Dim cError As String
Dim sToken As String

sToken = getTokenMall(cError)
If cError <> "" Then
    MsgBox (cError)
Else
    MsgBox sToken
    Clipboard.Clear
    Clipboard.SetText sToken
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmdSave_Click()
myreplace
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "baseUrl", addstring(xBaseUrl.text))
aInsert = AddFlag(aInsert, "AuthUrl", addstring(xAuthUrl.text))
aInsert = AddFlag(aInsert, "invoiceUrl", addstring(xInvoiceUrl.text))
aInsert = AddFlag(aInsert, "mail", addstring(xMail.text))
aInsert = AddFlag(aInsert, "Password", addstring(xPassword.text))
On Error GoTo myError
If IsEmpty(myField("Select id from SettingMall", con)) Then
   con.Execute addInsert(aInsert, "SettingMall")
Else
   con.Execute addUpdate(aInsert, "SettingMall", "Id = 1")
End If
Inform " „ «· ⁄œÌ·"
myreplace = True
Exit Function
myError:
MsgBox Err.Description
Err.Clear
End Function

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con
myload
cmdSave.Enabled = bSupermode
cmdGetToken.Visible = bSupermode
End Sub
Private Sub myload()
Dim loctable As New ADODB.Recordset
Set loctable = cmd("Select * from SettingMall", con).Execute
If Not loctable.EOF Then
    xBaseUrl.text = loctable!baseUrl & ""
    xAuthUrl.text = loctable!authUrl & ""
    xInvoiceUrl.text = loctable!invoiceUrl & ""
    xMail.text = loctable!mail & ""
    xPassword.text = loctable!PassWord & ""
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
DefineUrl con
closeCon con
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyAscii = 0
    End If
ElseIf KeyAscii = 19 And cmdSave.Enabled Then
    cmdSave_Click
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub xPasswordUrl_GotFocus()
myGotFocus xPasswordUrl
End Sub
Private Sub xPasswordUrl_LostFocus()
myLostFocus xPasswordUrl
End Sub
Private Sub xuserNameUrl_GotFocus()
myGotFocus xuserNameUrl
End Sub
Private Sub xuserNameUrl_LostFocus()
myLostFocus xuserNameUrl
End Sub
Private Sub xInvoiceUrl_GotFocus()
myGotFocus xInvoiceUrl
End Sub
Private Sub xInvoiceUrl_LostFocus()
myLostFocus xInvoiceUrl
End Sub
Private Sub xBaseUrl_GotFocus()
myGotFocus xBaseUrl
End Sub
Private Sub xBaseUrl_LostFocus()
myLostFocus xBaseUrl
End Sub
Private Sub xAuthUrl_GotFocus()
myGotFocus xAuthUrl
End Sub
Private Sub xAuthUrl_LostFocus()
myLostFocus xAuthUrl
End Sub
