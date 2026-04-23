Attribute VB_Name = "Module2"
Public Declare Function MessageBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" ( _
    ByVal hwnd As Long, _
    ByVal lpText As String, _
    ByVal lpCaption As String, _
    ByVal uType As Long, _
    ByVal wLanguageID As Long, _
    ByVal dwMilliseconds As Long) As Long




