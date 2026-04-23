VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpSup11 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
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
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2205
      Width           =   6180
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "œ›⁄«  ‰ﬁœÌ…"
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
         Height          =   420
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   180
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "‘Ìﬂ« "
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
         Height          =   420
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "ﬁÌÊœ Õ”«»« "
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
         Height          =   420
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " ”ÊÌ« "
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
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   180
         Value           =   1  'Checked
         Width           =   1140
      End
   End
   Begin VB.CommandButton CmdApply 
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
      Left            =   4680
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   2925
      Width           =   1500
   End
   Begin VB.CommandButton cmdClear 
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
      Left            =   3150
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   2925
      Width           =   1500
   End
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1620
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   2925
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   6180
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " ﬁ”Ì„ «·œ›⁄«  ··„Ê—œÌ‰"
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
         Height          =   240
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1755
         Width           =   3210
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   1230
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   990
         Width           =   1680
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xmosm 
         Height          =   330
         Left            =   1305
         TabIndex        =   3
         Top             =   1350
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
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
      Begin VB.Label Label5 
         Caption         =   "„Ê”„ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4845
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1395
         Width           =   1005
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   270
         Width           =   3345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4905
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   315
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4905
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   720
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
Attribute VB_Name = "rpSup11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearch As New Search_abd

Private Sub cmdClear_Click()
'DefineText Me
xdate1.text = ""
xDate2.text = ""
xCode.text = ""
xCodeDesca.Caption = ""
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub

Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xmosm_GotFocus()
myGotFocus xmosm
End Sub
Private Sub xmosm_LostFocus()
myLostFocus xmosm
If Not xmosm.MatchedWithList Then xmosm.BoundText = ""
End Sub
Sub Form_Load()
FixRpImage Me

openCon con

Set data4.Recordset = cmd("Select MOSM,DescA From MOSM     ORDER by DATE ", con).Execute
Set xmosm.RowSource = data4
xmosm.ListField = "Desca"
xmosm.BoundColumn = "MOSM"
xmosm.BoundText = cPMosm
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate2.text) And Trim(xDate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(3), cHead2 As String
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim loctable As ADODB.Recordset
Dim cString As String
Dim cWhere As String

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable


'cHead2 = "  ﬁ—Ì— „Ê—œÌ‰ "
'
'If Check2.Value = 1 Then
'    'cString = " SELECT  * FROM vw_SUP_PAY"
'    cString = "SELECT code,suppdesca,doc_no,[date],[value],DESCA,mosm,box,type_Desca " & _
'              " From vw_SUP_PAY"
'    cHead2 = cHead2 & " ‰ﬁœÌ…  "
'
'    If xmosm.MatchedWithList Then
'        cWhere = " AND  mosm = " & MyParn(xmosm.BoundText)
'        aHeader(0) = "[" & "„Ê”„ : " & xmosm.BoundText & "]"
'    End If
'
'    If Trim(xCode.text) <> "" Then
'        cWhere = cWhere & " AND CODE = " & MyParn(xCode.text)
'        aHeader(2) = "[" & "··„Ê—œ : " & xCodeDesca.Caption & "]"
'    End If
'
'
'    If IsDate(xdate1.text) Then
'        cWhere = cWhere & " AND date >= " & DateSq(xdate1.text)
'        aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
'    End If
'
'    If IsDate(xDate2.text) Then
'        cWhere = cWhere & " AND date <= " & DateSq(xDate2.text)
'        aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
'    End If
'
'    cWhere = Mid(cWhere, 6)
'    If cWhere <> "" Then
'        cString = cString & " WHERE " & cWhere
'    End If
'End If
''''' chq
'
'If Check3.Value = 1 Then
'    cWhere = ""
'    If cString <> "" Then
'        cString = cString & " UNION ALL "
'    End If
'
'    'cString = cString & "SELECT  *, ' ‘Ìﬂ ' FROM Q_chq"
'    cString = cString & _
'            "SELECT [CODE1]" & _
'            ",[SUPPDESCA]" & _
'            ",[ser_NO]" & _
'             ",[DATE_r]" & _
'            ",[VALUE]" & _
'            ",[DESCA]" & _
'            ",[MOSM]" & _
'            ",[BOX]" & _
'            "FROM Q_CHQ"
'
'    cHead2 = cHead2 & " ‘Ìﬂ«  "
'
'    If xmosm.MatchedWithList Then
'        cWhere = " AND mosm = " & MyParn(xmosm.BoundText)
'    End If
'
'    If Trim(xCode.text) <> "" Then
'        cWhere = cWhere & " and CODE1 = " & MyParn(xCode.text)
'    End If
'
'
'    If IsDate(xdate1.text) Then
'        cWhere = cWhere & " and date_r >= " & DateSq(xdate1.text)
'    End If
'
'    If IsDate(xDate2.text) Then
'        cWhere = cWhere & " and date_r <= " & DateSq(xDate2.text)
'    End If
'
'    cWhere = Mid(cWhere, 6)
'    If cWhere <> "" Then
'        cString = cString & " WHERE " & cWhere
'    End If
'End If
''''' ﬁÌÊœ
'
'If Check4.Value = 1 Then
'    cWhere = ""
'    If cString <> "" Then
'        cString = cString & " UNION ALL "
'    End If
'
'    cString = cString & " SELECT   file4_11.code ," & _
'             " file4_10.desca , doc_id , date , (pay-sal) , file4_11.desca , mosm , '' , ' ﬁÌœ ' as cashtype FROM file4_11 inner join file4_10 on file4_10.code = file4_11.code" & _
'             " where type = 'T' "
'
'    cHead2 = cHead2 & " ﬁÌÊœ "
'
'    If xmosm.MatchedWithList Then
'        cWhere = cWhere & " and mosm = " & MyParn(xmosm.BoundText)
'        aHeader(0) = "[" & "„Ê”„ : " & xmosm.BoundText & "]"
'    End If
'
'    If Trim(xCode.text) <> "" Then
'        cWhere = cWhere & " and file4_11.CODE = " & MyParn(xCode.text)
'        aHeader(2) = "[" & "··„Ê—œ : " & xCodeDesca.Caption & "]"
'    End If
'
'
'    If IsDate(xdate1.text) Then
'        cWhere = cWhere & " and date >= " & DateSq(xdate1.text)
'        aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
'    End If
'
'    If IsDate(xDate2.text) Then
'        cWhere = cWhere & " and date <= " & DateSq(xDate2.text)
'        aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
'    End If
'
'    cString = cString & cWhere
'End If
'
'If Check5.Value Then
'    cWhere = ""
'    If cString <> "" Then
'        cString = cString & " UNION ALL "
'    End If
'
'    cString = cString & _
'            "SELECT  code ,SUPPDESCA , doc_no , date , value , desca , mosm , box  , '  ”ÊÌ…  ' as cashtype" & _
'            " FROM Q_CASH0"
'
'    cHead2 = cHead2 & "  ”ÊÌ«  "
'
'    If xmosm.MatchedWithList Then
'        cWhere = " and mosm = " & MyParn(xmosm.BoundText)
'        aHeader(0) = "[" & "„Ê”„ : " & xmosm.BoundText & "]"
'    End If
'
'    If Trim(xCode.text) <> "" Then
'        cWhere = cWhere & " and CODE = " & MyParn(xCode.text)
'        aHeader(2) = "[" & "··„Ê—œ : " & xCodeDesca.Caption & "]"
'    End If
'
'    If IsDate(xdate1.text) Then
'        cWhere = cWhere & " and date >= " & DateSq(xdate1.text)
'        aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
'    End If
'
'    If IsDate(xDate2.text) Then
'        cWhere = cWhere & " and date <= " & DateSq(xDate2.text)
'        aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
'    End If
'
'    cWhere = Mid(cWhere, 6)
'    If cWhere <> "" Then
'        cString = cString & " WHERE " & cWhere
'    End If
'End If
'
Dim aPrm As Variant
If xmosm.MatchedWithList Then
    aPrm = AddFlag(aPrm, "MOSM", xmosm.BoundText)
    aHeader(0) = "„Ê”„ : " & xmosm.BoundText
End If

If Trim(xCode.text) <> "" Then
    aPrm = AddFlag(aPrm, "CODE", xCode.text)
    aHeader(1) = "··„Ê—œ : " & xCodeDesca.Caption
End If
 
If IsDate(xdate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xdate1.text))
    aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
End If

If IsDate(xDate2.text) Then
    aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xDate2.text))
    aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
End If

If Check2.Value = 1 Then
    aPrm = AddFlag(aPrm, "ISCASH", "1")
    aHeader(3) = Check2.Caption
End If

If Check3.Value = 1 Then
    aPrm = AddFlag(aPrm, "ISCHQ", "1")
    aHeader(3) = aHeader(3) & Tr(aHeader(3), "-") & Check3.Caption
End If

If Check4.Value = 1 Then
    aPrm = AddFlag(aPrm, "ISENTRY", "1")
    aHeader(3) = aHeader(3) & Tr(aHeader(3), "-") & Check4.Caption
End If

If Check5.Value = 1 Then
    aPrm = AddFlag(aPrm, "ISSET", "1")
    aHeader(3) = aHeader(3) & Tr(aHeader(3), "-") & Check4.Caption
End If


With loctable
Set loctable = cmd("[rp].[sp_sup_pay]", con, adStoredProc, aPrm).Execute
Do Until loctable.EOF
    temptable.AddNew
    temptable!str11 = loctable!code
    temptable!str1 = loctable!suppdesca
    temptable!str2 = loctable!doc_no
    temptable!str3 = loctable!TYPE_dESCA & "  " & loctable!DESCA
    temptable!Date1 = loctable!Date

    temptable!VAL1 = loctable!Value
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 4))
    temptable!str5 = Me.Caption
   ' If IsDate(loctable!Check_Date) Then
'    temptable!str12 = Format(loctable!Check_Date, "DD-MM-YYYY")
'    temptable!str11 = TurnValue(loctable!Check_No, "", Null)
    
    temptable.Update
    loctable.MoveNext
Loop
End With
temptable.Requery
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If Check1.Value Then
    Main.REPORT1.ReportFileName = App.Path & "\Reports\Sup11.rpt"
Else
    Main.REPORT1.ReportFileName = App.Path & "\Reports\Sup11G.rpt"
End If
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
loctable.Close
temptable.Close
Set loctable = Nothing
Set temptable = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    suplookup Me, oSearch
End If
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
xCodeDesca.Caption = ""
If xCode.text = "" Then Exit Sub
xCodeDesca.Caption = myField("select desca from FILE4_10 where code = " & MyParn(xCode.text), con) & ""
End Sub
Sub myProc()
    xCode.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
End Sub

