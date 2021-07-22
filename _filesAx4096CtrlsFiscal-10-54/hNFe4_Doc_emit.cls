VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_emit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_CNPJ   As String
Public m_xNome As String
Public m_xFant As String
Public m_IE As String

Public m_IM As String
Public m_CNAE As String
Public m_CRT As String

Public m_enderEmit As New hNFe4_Doc_emit_enderEmit

