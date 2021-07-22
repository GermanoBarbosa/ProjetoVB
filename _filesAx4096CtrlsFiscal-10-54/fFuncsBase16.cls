VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fFuncsBase16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Decode(ByVal Text As String) As String
Dim cc As Long, Char As String * 2
    'Select Case Left$(Text, 2)
    'Case "00"
        For cc = 1 To Len(Text) Step 2
            Char = Mid$(Text, cc, 2)
            Decode = Decode & Chr$(Val("&H" & Char))
        Next
    'End Select
End Function

Public Function Encode(ByVal Text As String) As String
Dim cc As Long, Char As String * 2
    'Text = "00" & Text
    For cc = 1 To Len(Text)
        Char = Mid$(Text, cc, 1)
        Encode = Encode & Right("00" & Hex(Asc(Char)), 2)
    Next
End Function

 

