VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ship_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ???? ??????: clsUser

' ??????? ???? ???? ?????
Private m_UserID As Long
Private m_UserName As String

' ????? ???? ???????? (ID)
Public Property Get UserID() As Long
    UserID = m_UserID
End Property

Public Property Let UserID(ByVal Value As Long)
    m_UserID = Value
End Property

' ????? ??? ???????? (Name)
Public Property Get UserName() As String
    UserName = m_UserName
End Property

Public Property Let UserName(ByVal Value As String)
    m_UserName = Value
End Property
Public Sub PopulateFromRecordset(ByRef rs As ADODB.Recordset)
    If Not rs.EOF Then
        m_UserID = rs.Fields("UserID").Value
        m_UserName = rs.Fields("UserName").Value
    End If
End Sub

