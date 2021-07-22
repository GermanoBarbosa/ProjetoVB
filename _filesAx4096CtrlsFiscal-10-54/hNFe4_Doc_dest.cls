VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_dest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'indIEDest
'1=Contribuinte ICMS (informar a IE do destinatário);
'2=Contribuinte isento de Inscrição no cadastro de Contribuintes

Public m_CNPJ As String
Public m_CPF As String
Public m_xNome As String
Public m_indIEDest As String
Public m_IE As String
Public m_email As String

Public m_enderDest As hNFe4_Doc_dest_enderDest

