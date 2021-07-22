VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_dest_enderDest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_xLgr As String
Public m_nro As String
Public m_xBairro As String
Public m_cMun As String
Public m_xMun As String
Public m_UF As String
Public m_CEP As String
Public m_cPais As String
Public m_xPais As String
Public m_fone As String
Public m_use As String

Private Sub Class_Initialize()
    m_cPais = 1058
    m_xPais = "BRASIL"
End Sub

