Attribute VB_Name = "mFuncs"
Option Explicit

Public Base16 As New fFuncsBase16

Public m_path As String

Public Function FileExists(sFile As String) As Boolean
    On Error Resume Next
    If GetAttr(sFile) = -1 Then
    Else
        FileExists = True
    End If
    On Error GoTo 0
End Function

Function SetFileBytes(ByVal FileName As String, mFileBytes As String) As String
Dim NumFile As Long
    NumFile = FreeFile
    If FileExists(FileName) Then
        Kill FileName
    End If
    Open FileName For Binary As NumFile
    Put NumFile, , mFileBytes
    Close NumFile
End Function

Function ClearXML(ByVal mXML As String) As String
Dim txt_out As String
    txt_out = mXML
    Do While InStr(txt_out, "  ") > 0
        txt_out = Replace(txt_out, "  ", "|")
    Loop
    txt_out = Replace(txt_out, "|", "")
    txt_out = Replace(txt_out, vbTab, "")
    'txt_out = Replace(txt_out, vbTab, "*")
    'txt_out = Replace(txt_out, "> <", "><")
    txt_out = Replace(txt_out, vbCr, "")
    txt_out = Replace(txt_out, vbLf, "")
    ClearXML = txt_out
End Function

Public Function IsIDE() As Boolean
Static mIsIDE As Boolean
    Debug.Assert SetTrue(mIsIDE) = False
    IsIDE = mIsIDE
End Function

Function SetTrue(mBool As Boolean) As Boolean
    mBool = True
End Function


Function GetFileBytes(ByVal FileName As String, Optional m_Len As Long = -1) As String
    Dim NumFile As Long
    NumFile = FreeFile
    Open FileName For Binary As NumFile
    If m_Len <= -1 Then
        m_Len = LOF(NumFile)
    End If
    GetFileBytes = Space(m_Len)
    Get NumFile, , GetFileBytes
    Close NumFile
End Function
