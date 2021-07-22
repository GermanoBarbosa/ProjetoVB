VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_det_imposto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_vTotTrib As Double     'Valor aproximado total de tributos federais, estaduais e municipais.

Public m_ICMS_ICMS00_orig
Public m_ICMS_ICMS00_CST  As String
Public m_ICMS_ICMS00_modBC
Public m_ICMS_ICMS00_vBC    As Double
Public m_ICMS_ICMS00_pICMS As Double
Public m_ICMS_ICMS00_vICMS As Double

Public m_pST As Double
Public m_vICMSSTRet As Double
Public m_ICMS_ICMS500_vBCST As Double


Public m_IPI_cEnq
Public m_IPI_IPINT_CST  As String
Public m_PIS_PISAliq_CST  As String
Public m_PIS_PISAliq_vBC As Double
Public m_PIS_PISAliq_pPIS As Double
Public m_PIS_PISAliq_vPIS As Double

Public m_COFINS_COFINSAliq_CST As String
Public m_COFINS_COFINSAliq_vBC As Double
Public m_COFINS_COFINSAliq_pCOFINS As Double
Public m_COFINS_COFINSAliq_vCOFINS As Double

Public m_ICMSUFDest_vBCUFDest
Public m_ICMSUFDest_pFCPUFDest
Public m_ICMSUFDest_pICMSUFDest
Public m_ICMSUFDest_pICMSInter
Public m_ICMSUFDest_pICMSInterPart
Public m_ICMSUFDest_vFCPUFDest
Public m_ICMSUFDest_vICMSUFDest
Public m_ICMSUFDest_vICMSUFRemet

