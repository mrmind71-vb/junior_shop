VERSION 5.00
Begin VB.Form ApiStockSettingfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "÷»ÿ «⁄œ«œ   ÿ»Ìﬁ «·Ã—œ"
   ClientHeight    =   4485
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3150
      Width           =   11445
      Begin VB.TextBox xSize 
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
         Left            =   1845
         MaxLength       =   100
         TabIndex        =   1
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "⁄œœ «·”Ã·«  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.CommandButton CmdExit 
      Height          =   510
      Left            =   1395
      MaskColor       =   &H00FFFFFF&
      Picture         =   "MobilSetting.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3870
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin VB.CommandButton cmdSave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      Picture         =   "MobilSetting.frx":4AB6
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Õ›Ÿ"
      Top             =   3870
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   3120
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   11445
      Begin VB.TextBox xuserPassword 
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
         Left            =   2025
         MaxLength       =   200
         TabIndex        =   21
         Top             =   2700
         Width           =   9285
      End
      Begin VB.TextBox xusername 
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
         Left            =   2025
         MaxLength       =   200
         TabIndex        =   20
         Top             =   2340
         Width           =   9285
      End
      Begin VB.TextBox xMainUrl 
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
         Left            =   2025
         MaxLength       =   200
         TabIndex        =   19
         Top             =   180
         Width           =   9285
      End
      Begin VB.TextBox xitemsUrl 
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
         Left            =   2025
         MaxLength       =   200
         TabIndex        =   18
         Top             =   540
         Width           =   9285
      End
      Begin VB.TextBox xBranchURL 
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
         Left            =   2025
         MaxLength       =   200
         TabIndex        =   17
         Top             =   900
         Width           =   9285
      End
      Begin VB.TextBox xStockDocURL 
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
         Left            =   2025
         MaxLength       =   200
         TabIndex        =   16
         Top             =   1260
         Width           =   9285
      End
      Begin VB.TextBox xStockItemsURL 
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
         Left            =   2025
         MaxLength       =   200
         TabIndex        =   15
         Top             =   1620
         Width           =   9285
      End
      Begin VB.TextBox xTokenURL 
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
         Left            =   2025
         MaxLength       =   200
         TabIndex        =   0
         Top             =   1980
         Width           =   9285
      End
      Begin VB.Label Label8 
         Caption         =   "Tokern Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   14
         Top             =   2745
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "Tokern User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   13
         Top             =   2385
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Token URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   2025
         Width           =   1545
      End
      Begin VB.Label Label7 
         Caption         =   "Stock Items URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   11
         Top             =   1665
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Stock Doc URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   9
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Branch URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   8
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "Items Url"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   7
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label10 
         Caption         =   "Main Url"
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
         TabIndex        =   5
         Top             =   270
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ApiStockSettingfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdSave_Click()
myreplace
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "MainUrl", addstring(xMainUrl.text))
aInsert = AddFlag(aInsert, "itemsURL", addstring(xitemsUrl.text))
aInsert = AddFlag(aInsert, "branchUrl", addstring(xBranchURL.text))
aInsert = AddFlag(aInsert, "stockDocURL", addstring(xStockDocURL.text))
aInsert = AddFlag(aInsert, "stockItemsURL", addstring(xStockItemsURL.text))
aInsert = AddFlag(aInsert, "TokenURL", addstring(xTokenURL.text))
aInsert = AddFlag(aInsert, "username", addstring(xusername.text))
aInsert = AddFlag(aInsert, "userpassword", addstring(xuserPassword.text))
aInsert = AddFlag(aInsert, "Size", mRound(xSize.text))

On Error GoTo myError
con.BeginTrans
If IsEmpty(myField("Select id from SettingApiStock", con)) Then
   aInsert = AddFlag(aInsert, "ID", "1")
   con.Execute addInsert(aInsert, "SettingApiStock")
Else
   con.Execute addUpdate(aInsert, "SettingApiStock", "Id = 1")
End If
con.CommitTrans
Inform " „ «· ⁄œÌ·"
myreplace = True
Exit Function
myError:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
'cmdTrans.Visible = retFlag(aSec, "manager")

Handlecontrols
openCon con
myload
End Sub
Private Sub myload()
Dim loctable As New ADODB.Recordset
Set loctable = cmd("Select * from SettingApiStock", con).Execute
If Not loctable.EOF Then
    xMainUrl.text = loctable!mainUrl & ""
    xitemsUrl.text = loctable!itemsURL & ""
    xTokenURL.text = loctable!TokenUrl & ""
    xStockDocURL.text = loctable!StockDocURL & ""
    xStockItemsURL.text = loctable!itemsURL & ""
    xusername.text = loctable!UserName & ""
    xuserPassword.text = loctable!userpassword & ""
    xSize.text = Myvalue(loctable!Size)
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'FixUrl con
closeCon con
Set ApiStockSettingfrm = Nothing
End Sub
Private Sub xToken_GotFocus()
myGotFocus xToken
End Sub
Private Sub xToken_LostFocus()
myLostFocus xToken
End Sub
Private Sub xSize_GotFocus()
myGotFocus xSize
End Sub
Private Sub xSize_LostFocus()
myLostFocus xSize
End Sub
Private Sub xCustomerUrl_GotFocus()
myGotFocus xCustomerUrl
End Sub
Private Sub xCustomerUrl_LostFocus()
myLostFocus xCustomerUrl
End Sub
Private Sub xUserUrl_GotFocus()
myGotFocus xUserUrl
End Sub
Private Sub xUserUrl_LostFocus()
myLostFocus xUserUrl
End Sub

Private Sub xOrderUrl_GotFocus()
myGotFocus xOrderUrl
End Sub
Private Sub xOrderUrl_LostFocus()
myLostFocus xOrderUrl
End Sub
Private Sub xitemUrl_GotFocus()
myGotFocus xitemUrl
End Sub
Private Sub xitemUrl_LostFocus()
myLostFocus xitemUrl
End Sub
Private Sub xModelUrl_GotFocus()
myGotFocus xModelUrl
End Sub
Private Sub xModelUrl_LostFocus()
myLostFocus xModelUrl
End Sub
Private Sub xSectionUrl_GotFocus()
myGotFocus xSectionUrl
End Sub
Private Sub xSectionUrl_LostFocus()
myLostFocus xSectionUrl
End Sub
Private Sub xMainUrl_GotFocus()
myGotFocus xMainUrl
End Sub
Private Sub xMainUrl_LostFocus()
myLostFocus xMainUrl
End Sub
Private Sub Handlecontrols()
'xMainUrl.Enabled = False
'xCustomerUrl.Enabled = False
'xSectionUrl.Enabled = False
'xModelUrl.Enabled = False
'xitemUrl.Enabled = False
'xOrderUrl.Enabled = False
'xToken.Enabled = False
End Sub
