VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form online_ship_headerfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "÷»ÿ „·› ”Õ» ”œ«œ „‰ Œ·«· «·„Êﬁ⁄"
   ClientHeight    =   3720
   ClientLeft      =   -15
   ClientTop       =   435
   ClientWidth     =   8430
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
   ScaleHeight     =   3720
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3030
      Left            =   45
      TabIndex        =   6
      Top             =   -45
      Width           =   8160
      Begin VB.TextBox xVAT_header 
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
         TabIndex        =   13
         Top             =   1845
         Width           =   5100
      End
      Begin VB.TextBox xCOD_header 
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
         TabIndex        =   11
         Top             =   1125
         Width           =   5100
      End
      Begin VB.TextBox xNet_Header 
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
         Top             =   2205
         Width           =   5100
      End
      Begin VB.TextBox xFees_header 
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
         Top             =   1485
         Width           =   5100
      End
      Begin VB.TextBox xorder_header 
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
         Top             =   765
         Width           =   5100
      End
      Begin MSDataListLib.DataCombo xShip 
         Height          =   330
         Left            =   2880
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   405
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "VAT Header"
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
         TabIndex        =   14
         Top             =   1890
         Width           =   2310
      End
      Begin VB.Label Label3 
         Caption         =   "COD Header"
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
         TabIndex        =   12
         Top             =   1170
         Width           =   2310
      End
      Begin VB.Label Label2 
         Caption         =   "Shipping Company Name"
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
         TabIndex        =   10
         Top             =   450
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Net Value Header"
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
         Top             =   2250
         Width           =   1950
      End
      Begin VB.Label Label10 
         Caption         =   "Fees Header"
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
         TabIndex        =   8
         Top             =   1530
         Width           =   2310
      End
      Begin VB.Label Label11 
         Caption         =   "Order ID Header"
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
         TabIndex        =   7
         Top             =   810
         Width           =   2175
      End
   End
   Begin VB.CommandButton CmdExit 
      Height          =   555
      Left            =   6840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ship_header.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3060
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
      Height          =   555
      Left            =   5355
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ship_header.frx":4AB6
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Õ›Ÿ"
      Top             =   3060
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "online_ship_headerfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sShip As String
Dim con As New adodb.Connection
Private Sub cmdSave_Click()
If myreplace Then Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "order_header", addstring(xorder_header.text))
aInsert = AddFlag(aInsert, "COD_header", addstring(xCOD_header.text))
aInsert = AddFlag(aInsert, "fees_header", addstring(xFees_header.text))
aInsert = AddFlag(aInsert, "vat_header", addstring(xVAT_header.text))
aInsert = AddFlag(aInsert, "net_header", addstring(xNet_Header.text))
con.Execute addUpdate(aInsert, "SHIP", "CODE = " & MyParn(xShip.BoundText))
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con
Set DATA1.Recordset = myCmd("SELECT * FROM SHIP WHERE STOPED = 0 ORDER BY DESCA", con)
Set xShip.RowSource = DATA1
xShip.ListField = "Desca"
xShip.BoundColumn = "Code"
If sShip <> "" Then
    xShip.BoundText = sShip
    xShip.Enabled = False
    myload
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'DefineUrl con
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
Private Sub xShip_Click(Area As Integer)
myload
End Sub
Private Sub myload()
DATA1.Recordset.Filter = "CODE = " & xShip.BoundText
If Not DATA1.Recordset.EOF Then
    xorder_header.text = DATA1.Recordset!order_header & ""
    xCOD_header.text = DATA1.Recordset!cod_HEADER & ""
    xFees_header.text = DATA1.Recordset!fees_header & ""
    xVAT_header.text = DATA1.Recordset!vat_header & ""
    xNet_Header.text = DATA1.Recordset!net_header & ""
Else
    xorder_header.text = ""
    xCOD.text = ""
    xFees_header.text = ""
    xVAT_header.text = ""
    xNet_Header.text = ""
End If
DATA1.Recordset.Filter = ""
End Sub
