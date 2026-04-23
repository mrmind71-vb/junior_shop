VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form send_imagesfrm 
   Caption         =   "Send Images"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Images"
      Height          =   600
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4470
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   1
      Top             =   945
      Visible         =   0   'False
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "send_imagesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADOdb.Connection
Dim con_image As New ADOdb.Connection
Private Sub cmdSend2_Click()
Dim fs As New FileSystemObject

Dim rs As ADOdb.Recordset
Dim mstream As ADOdb.Stream
Dim imageID As Integer


Set mstream = New ADOdb.Stream
mstream.Type = adTypeBinary ' Set the stream type to binary data
mstream.Open

'myPath = App.Path & "\DOCUMENTS\" ' Ensure path ends with a backslash

'sTemp = Dir(myPath & FileExtension)
 
Dim Folder As Folder
Set Folder = fs.GetFolder(App.Path & "\Documents")
 
' Set rs = New ADODB.Recordset
'rs.Open "SELECT * FROM ACCOUNT_IMAGES", con, adOpenKeyset, adLockOptimistic

'nCount = CountFilesWithExt(myPath, "jpg")

Dim command As New ADOdb.command

        ' --- Load Image into ADODB.Stream ---
        
Dim sCaption As String
Dim sDoc_no As String
Dim File As File
prog1.Visible = True

sCaption = Me.Caption
For Each File In Folder.Files
    nCount = nCount + 1
    prog1.Value = Round(nCount / Folder.Files.Count, 2) * 100
    Me.Caption = nCount & " From " & Folder.Files.Count
    If LCase(fs.GetExtensionName(File.Path)) = "jpg" Then
        mstream.LoadFromFile File.Path
        
        sAll = Mid(File.Name, 1, Len(File.Name) - 4)
        sid_image = Mid(sAll, (InStrRev(sAll, "-") + 1))
        
        sAll_No_Id = Mid(sAll, 1, (InStrRev(sAll, "-") - 1))
        
        sDoc_no = Mid(sAll_No_Id, 1, (InStrRev(sAll_No_Id, "-") - 1))
        
        sId_cash = Mid(sAll_No_Id, (InStrRev(sAll_No_Id, "-") + 1))
        
        aPrm = AddFlag(Empty, "id_image", sid_image)
        aPrm = AddFlag(aPrm, "id_cash", sId_cash)
        aPrm = AddFlag(aPrm, "doc_no", sDoc_no)
        aPrm = AddFlag(aPrm, "image", mstream.Read)
        
        Set command = cmd("[dbo].[sp_add_account_image]", con, adStoredProc, aPrm)
        command.Execute
    End If
Next
Me.Caption = sCaption
'    ' *** REPLACE WITH YOUR VALUES ***
'
'    Do Until loctable.EOF
'
'    Dim imagePath As String
'    imagePath = "C:\path\to\your\image.jpg" ' e.g., "C:\temp\myphoto.jpg"
'
'    Dim tableName As String
'    tableName = "YourImageTable"
'    Dim fieldName As String
'    fieldName = "ImageField"
'    Dim imageID As Integer
'    imageID = 1 ' Or use an ID from an existing record
'
'    openCon con
'    ' Ensure connection is open
'    Set mstream = New ADODB.Stream
'    mstream.Type = adTypeBinary ' Set the stream type to binary data
'    mstream.Open
'    mstream.LoadFromFile imagePath  ' Load the image file into the stream
'
'    Set rs = New ADODB.Recordset
'    ' Open the recordset with adLockOptimistic to allow updates
'    rs.Open "SELECT * FROM " & tableName & " WHERE ID = " & imageID, cn, adOpenKeyset, adLockOptimistic
'
'    If Not rs.EOF Then
'        ' Update the existing record
'        rs.Fields(fieldName).Value = mstream.Read ' Read binary data from stream and assign to the field
'        rs.Update
'        MsgBox "Image updated successfully!"
'    Else
'        ' Optional: Code to Add a new record (if needed)
'        ' rs.AddNew
'        ' rs!ID = imageID
'        ' rs.Fields(fieldName).Value = mstream.Read
'        ' rs.Update
'        ' MsgBox "New image record added successfully!"
'    End If
'
'    ' Cleanup
'    mstream.Close
'    rs.Close
'    Set mstream = Nothing
'    Set rs = Nothing
    ' Keep connection open or close as needed by your application
End Sub
Private Sub cmdsend_Click()
Dim fs As New FileSystemObject

Dim rs As ADOdb.Recordset
Dim mstream As ADOdb.Stream
Dim imageID As Integer


Set mstream = New ADOdb.Stream
mstream.Type = adTypeBinary ' Set the stream type to binary data
mstream.Open

 
' Set rs = New ADODB.Recordset
'rs.Open "SELECT * FROM ACCOUNT_IMAGES", con, adOpenKeyset, adLockOptimistic

'nCount = CountFilesWithExt(myPath, "jpg")

Dim loctable As New ADOdb.Recordset
Set loctable = cmd("SELECT * FROM ACCOUNT_P WHERE POSTED = 0", con).Execute


Dim command As New ADOdb.command

        ' --- Load Image into ADODB.Stream ---
        
Dim sCaption As String
Dim sDoc_no As String
Dim File As File
Dim nCount As Long

prog1.Visible = True

sCaption = Me.Caption
Do Until loctable.EOF
    nCount = nCount + 1
    prog1.Value = Round(nCount / loctable.RecordCount, 2) * 100
    
    Me.Caption = nCount & " From " & loctable.RecordCount
    
    cFile = App.Path & "\Documents\" & loctable!photo & ".jpg"
    
    If fs.FileExists(cFile) Then
    
        mstream.LoadFromFile cFile
        
        'sAll = Mid(file.Name, 1, Len(file.Name) - 4)
'        sid_image = loctable!ID
'
'        sDoc_no = Mid(sAll_No_Id, 1, (InStrRev(sAll_No_Id, "-") - 1))
'
'        sId_cash = Mid(sAll_No_Id, (InStrRev(sAll_No_Id, "-") + 1))
        
        aPrm = AddFlag(Empty, "id_image", loctable!ID)
        aPrm = AddFlag(aPrm, "id_cash", loctable!id_cash)
        aPrm = AddFlag(aPrm, "doc_no", loctable!doc_no)
        aPrm = AddFlag(aPrm, "image", mstream.Read)
        
        Set command = cmd("[dbo].[sp_add_account_image]", con_image, adStoredProc, aPrm)
        command.Execute
        
        con.Execute "UPDATE ACCOUNT_P SET POSTED = 1 WHERE ID = " & loctable!ID
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
Me.Caption = sCaption
End Sub
Private Sub Form_Load()
openCon con
openConDoc con_image
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con_image
closeCon con
End Sub

