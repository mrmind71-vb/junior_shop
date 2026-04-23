VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form cost_fixfrm 
   Caption         =   "÷»ÿ  ﬂ·›…"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAllItems 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "ﬂ· «·«’‰«›"
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
      Left            =   3060
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1710
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   1260
      MaskColor       =   &H00FFFFFF&
      Picture         =   "cost_fix.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   " ﬂ·›… «·„»Ì⁄« "
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3105
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1770
   End
   Begin MSComctlLib.StatusBar bar1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   7
      Top             =   2205
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   3570
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   2265
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   2265
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   780
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   8
      Top             =   2055
      Visible         =   0   'False
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "cost_fixfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nMode As Integer
Public nFlag As Integer
Dim con As New ADODB.Connection
Private Sub CmdApply_Click()
If nFlag = 0 Then
    fixPurchaseCostDates
    fixItemsCostDates
ElseIf nFlag = 1 Then
    fixSalesCostDates
ElseIf nFlag = 2 Then
    fixstockCostDates
ElseIf nFlag = 3 Then
    fixDamageCostDates
ElseIf nFlag = -1 Then
    fixPurchaseCostDates
    fixItemsCostDates True
    fixSalesCostDates
    fixstockCostDates
    fixDamageCostDates
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
'xdate1.text = myFormat_p(RetSetting("DATE1", TempSave(Me, nFlag & "")))
'XDATE2.text = myFormat_p(RetSetting("DATE2", TempSave(Me, nFlag & "")))

xDate1.text = myFormat_p(Year(Date) & "-01-01")
xdate2.text = myFormat_p(Date)
openCon con
If nFlag = 0 Then
   cmdApply.Caption = " ﬂ·›… «·„‘ —Ì« "
   chkAllItems.Visible = True
ElseIf nFlag = 1 Then
   cmdApply.Caption = " ﬂ·›… «·„»Ì⁄« "
ElseIf nFlag = 2 Then
   cmdApply.Caption = " ﬂ·›… «·Ã—œ"
ElseIf nFlag = 3 Then
   cmdApply.Caption = " ﬂ·›… «· «·›"
ElseIf nFlag = -1 Then
   cmdApply.Caption = "÷»ÿ «· ﬂ·›…"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
addSetting "DATE1", xDate1.text, TempSave(Me, nFlag & "")
addSetting "DATE2", xdate2.text, TempSave(Me, nFlag & "")
closeCon con
Set closefrm = Nothing
End Sub

Private Sub xdate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xdate2
End Sub


Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate xdate2
End Sub
Private Sub fixSalesCostDates()
'Dim nRecords As Long
'nRecords = fixCostSales(con, , , myFormat(xDate1.text), myFormat(xdate2.text), , , prog1, Me)
''nRecords = fixCostSalesPeriod(con, myFormat(xDate1.text), myFormat(xdate2.text), prog1, Me)
'Inform " „ ÷»ÿ  ﬂ·›…" & nRecords & " ”Ã·" & " „»Ì⁄« "
'nRecords = fixCostSalesRet(con, , , myFormat(xDate1.text), myFormat(xdate2.text), , , prog1, Me)
'Inform " „ ÷»ÿ  ﬂ·›…" & nRecords & " ”Ã·" & " „—œÊœ „»Ì⁄« "
'bar1.Panels(1).text = "⁄œœ «·”Ã·«  " & nRecords
End Sub
Private Sub fixPurchaseCostDates()
'Dim nRecords As Long
'nRecords = fixCostPurchase(con, , , , myFormat(xDate1.text), myFormat(xdate2.text), , prog1, Me)
'Inform " „ ÷»ÿ  ﬂ·›…" & nRecords & " ”Ã·" & " „‘ —Ì« "
'bar1.Panels(1).text = "⁄œœ «·”Ã·«  " & nRecords
End Sub
Private Sub fixstockCostDates()
Dim sError As String
Dim nRecords As Long

If Not fixCostStock(con, , , myFormat(xDate1.text), myFormat(xdate2.text), nRecords, sError) Then
   MsgBox sError
   Exit Sub
End If
Inform " „ ÷»ÿ  ﬂ·›…" & nRecords & " ”Ã·" & " Ã—œ"
bar1.Panels(1).text = "⁄œœ «·”Ã·«  " & nRecords
End Sub
Private Sub fixDamageCostDates()
'Dim nRecords As Long
'nRecords = fixCostDamage(con, , , myFormat(xDate1.text), myFormat(xdate2.text), , , prog1, Me)
'Inform " „ ÷»ÿ  ﬂ·›…" & nRecords & " ”Ã·" & " „·› Â«·ﬂ"
'bar1.Panels(1).text = "⁄œœ «·”Ã·«  " & nRecords
End Sub
Private Sub fixItemsCostDates(Optional bAll As Boolean = False)
'Dim nRecords As Long
'If chkAllItems.Value = 1 Or bAll Then
'    nRecords = fixCostItemGroup(con, , , , , , , prog1, Me, 100)
'Else
'    nRecords = fixCostItem(con, , , , , myFormat(xDate1.text), myFormat(xdate2.text), prog1, Me)
'End If
'Inform " „ ÷»ÿ  ﬂ·›…" & nRecords & " ”Ã·" & " ’‰›"
'bar1.Panels(1).text = "⁄œœ «·”Ã·«  " & nRecords
End Sub


