VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    

'https://desenvolvedores.migrate.com.br/downloads/

 
Public m_infAdic As New hNFe4_Doc_infAdic
Public m_ide As New hNFe4_Doc_ide
Public m_emit As New hNFe4_Doc_emit
Public m_dest As New hNFe4_Doc_dest
Public m_total As New hNFe4_Doc_total
Public m_Chave_NF As String
Public m_Ambiente As Long                   'Identificação de Ambiente: 1 - Produção, 2 - Homologação.
Public m_Endereco_QRCode As String
Public m_CSC As String

Dim m_Itens() As hNFe4_Doc_det, m_cItens As Long


Sub Clear_det()
    m_cItens = -1
    Erase m_Itens
    m_total.Clear
End Sub

Sub Add_det(m_Item As hNFe4_Doc_det, Optional m_flag As Long)
    
    
    If m_cItens = -1 And m_Ambiente = 2 And m_flag = 0 Then
        Dim m_Item2 As New hNFe4_Doc_det
        
        Set m_Item2 = m_Item
        With m_Item
            .prod.m_cProd = "1"
            .prod.m_xProd = "NOTA FISCAL EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
            '.prod.m_uCom = 1
            '.prod.m_qCom = 1
            '.prod.m_indTot = 1
            '.prod.m_vProd = 1
            '.prod.m_vUnCom = 1
            '.prod.m_vUnTrib = 1
            '.imposto.m_vTotTrib = 1
        End With
        Add_det m_Item2, 1
        Exit Sub
    End If
    
    m_cItens = m_cItens + 1
    ReDim Preserve m_Itens(m_cItens)
    Set m_Itens(m_cItens) = m_Item
    
    m_total.m_vBC = m_total.m_vBC + m_Item.imposto.m_ICMS_ICMS00_vBC
    m_total.m_vICMS = m_total.m_vICMS + m_Item.imposto.m_ICMS_ICMS00_vICMS
    m_total.m_vProd = m_total.m_vProd + m_Item.prod.m_vProd
    m_total.m_vFrete = m_total.m_vFrete + m_Item.prod.m_vFrete
    m_total.m_vPIS = m_total.m_vPIS + m_Item.imposto.m_PIS_PISAliq_vPIS
    m_total.m_vCOFINS = m_total.m_vCOFINS + m_Item.imposto.m_COFINS_COFINSAliq_vCOFINS
    m_total.m_vTotTrib = m_total.m_vTotTrib + m_Item.imposto.m_vTotTrib
    'm_total.m_vNF=m_total.m_vNF+
End Sub



Private Function Calculo_DV11(strNumero As String) As String
'declara as variáveis
    Dim intcontador, intnumero, intTotalNumero, intMultiplicador, intResto As Integer

    ' se nao for um valor numerico sai da função
    If Not IsNumeric(strNumero) Then
        Calculo_DV11 = ""
        Exit Function
    End If

    'inicia o multiplicador
    intMultiplicador = 9

    'pega cada caracter do numero a partir da direita
    For intcontador = Len(strNumero) To 1 Step -1

        'extrai o caracter e multiplica prlo multiplicador
        intnumero = Val(Mid(strNumero, intcontador, 1)) * intMultiplicador

        'soma o resultado para totalização
        intTotalNumero = intTotalNumero + intnumero

        'se o multiplicador for maior que 2 decrementa-o caso contrario atribuir valor padrao original
        intMultiplicador = IIf(intMultiplicador > 2, intMultiplicador - 1, 9)

    Next

    'calcula o resto da divisao do total por 11
    intResto = intTotalNumero Mod 11

    'verifica as exceções ( 0 -> DV=0    10 -> DV=X (para o BB) e retorna o DV
    Select Case intResto
    Case 0
        Calculo_DV11 = "0"
    Case 10
        Calculo_DV11 = "0"   'AQUI É X PARA BB
    Case Else
        Calculo_DV11 = Trim(Str(intResto))

    End Select

End Function

Function GetNota(m_Chave_2 As String, m_NF_NUM_NOTA As Long) As String
    Dim txt_xml   As New hStringBuilder
    Dim txt_xml2   As String
    Dim m_Chave   As New hStringBuilder
    Dim m_ChaveDV As String
    Dim cc        As Long
    Dim m_HashQRCode As String
    'txt_xml.Append "<?xml version=""1.0"" encoding=""UTF-8""?>"
    'txt_xml.Append "<nfeProc xmlns=""http://www.portalfiscal.inf.br/nfe"" versao=""4.00"">"

    '”NFe13 1403 11707347000195 65 003 0000004591064552496?
    'm_Chave.Append "NFe"
    m_NF_NUM_NOTA = m_ide.m_nNF
    m_Chave.Clear
    m_Chave.Append m_ide.m_cUF                              '02 - cUF - Código da UF do emitente do Documento Fiscal
    m_Chave.Append Format(m_ide.m_dhEmi, "YYMM")            '04 - AAMM - Ano e Mês de emissão da NF-e
    m_Chave.Append m_emit.m_CNPJ                            '14 - CNPJ - CNPJ do emitente
    m_Chave.Append m_ide.m_mod                              '02 - mod - Modelo do Documento Fiscal
    m_Chave.Append Format(Val(m_ide.m_serie), "000")        '03 - serie - Série do Documento Fiscal
    m_Chave.Append Format(Val(m_ide.m_nNF), "000000000")    '09 - nNF - Número do Documento Fiscal
    m_Chave.Append m_ide.m_tpEmis                           '01 - tpEmis – forma de emissão da NF-e
    m_Chave.Append Format(Val(m_ide.m_cNF), "00000000")     '08 - cNF - Código Numérico que compõe a Chave de Acesso
    m_ChaveDV = Calculo_DV11(m_Chave.toString)              '01 - cDV - Dígito Verificador da Chave de Acesso
    m_Chave.Append m_ChaveDV
    'MsgBox Len(m_Chave)
    m_Chave_2 = m_Chave
    m_Chave_NF = m_Chave
    
    txt_xml.Append Trim("   <NFe xmlns=""http://www.portalfiscal.inf.br/nfe"">")
    txt_xml.Append Trim("       <infNFe Id=""" & "NFe" & m_Chave & """ versao=""4.00"">")
    'txt_xml.Append Trim("       <infNFe Id=""NFe42161105865176000472550010000087601003965333"" versao=""4.00"">")
    txt_xml.Append Trim("           <ide>")
    txt_xml.Append Trim("               <cUF>" & m_ide.m_cUF & "</cUF>")
    txt_xml.Append Trim("               <cNF>" & m_ide.m_cNF & "</cNF>")
    txt_xml.Append Trim("               <natOp>VDA MERC ADQ TERCEIR</natOp>")
    'txt_xml.Append Trim("               <indPag>1</indPag>")
    txt_xml.Append Trim("               <mod>" & m_ide.m_mod & "</mod>")
    txt_xml.Append Trim("               <serie>" & m_ide.m_serie & "</serie>")
    txt_xml.Append Trim("               <nNF>" & m_ide.m_nNF & "</nNF>")
    txt_xml.Append Trim("               <dhEmi>" & cFormatDate_NF_1(m_ide.m_dhEmi) & "</dhEmi>")

    If m_ide.m_mod = "55" Then
        txt_xml.Append Trim("               <dhSaiEnt>" & cFormatDate_NF_1(m_ide.m_dhSaiEnt) & "</dhSaiEnt>")
    End If
    txt_xml.Append Trim("               <tpNF>1</tpNF>")
    txt_xml.Append Trim("               <idDest>1</idDest>")
    txt_xml.Append Trim("               <cMunFG>" & m_ide.m_cMunFG & "</cMunFG>")

    If m_ide.m_mod = "55" Then
        txt_xml.Append Trim("               <tpImp>1</tpImp>") '1=DANFE normal, Retrato;
    ElseIf m_ide.m_mod = "65" Then
        txt_xml.Append Trim("               <tpImp>4</tpImp>") '4=DANFE NFC-e;
    End If
    txt_xml.Append Trim("               <tpEmis>" & m_ide.m_tpEmis & "</tpEmis>")
    txt_xml.Append Trim("               <cDV>" & m_ChaveDV & "</cDV>")
    txt_xml.Append Trim("               <tpAmb>" & m_Ambiente & "</tpAmb>")
    txt_xml.Append Trim("               <finNFe>1</finNFe>")
    txt_xml.Append Trim("               <indFinal>1</indFinal>")

    If m_ide.m_mod = "55" Then
        txt_xml.Append Trim("               <indPres>9</indPres>")
    ElseIf m_ide.m_mod = "65" Then
        txt_xml.Append Trim("               <indPres>1</indPres>")
    End If
    txt_xml.Append Trim("               <procEmi>0</procEmi>")
    txt_xml.Append Trim("               <verProc>" & App.Major & "." & App.Minor & "." & App.Revision & "</verProc>")
    txt_xml.Append Trim("           </ide>")
    txt_xml.Append Trim("           <emit>")
    txt_xml.Append Trim("               <CNPJ>" & m_emit.m_CNPJ & "</CNPJ>")
    txt_xml.Append Trim("               <xNome>" & m_emit.m_xNome & "</xNome>")
    txt_xml.Append Trim("               <xFant>" & m_emit.m_xFant & "</xFant>")
    txt_xml.Append Trim("               <enderEmit>")
    txt_xml.Append Trim("                   <xLgr>" & m_emit.m_enderEmit.m_xLgr & "</xLgr>")
    txt_xml.Append Trim("                   <nro>" & m_emit.m_enderEmit.m_nro & "</nro>")
    txt_xml.Append Trim("                   <xBairro>" & m_emit.m_enderEmit.m_xBairro & "</xBairro>")
    txt_xml.Append Trim("                   <cMun>" & m_emit.m_enderEmit.m_cMun & "</cMun>")
    txt_xml.Append Trim("                   <xMun>" & m_emit.m_enderEmit.m_xMun & "</xMun>")
    txt_xml.Append Trim("                   <UF>" & m_emit.m_enderEmit.m_UF & "</UF>")
    txt_xml.Append Trim("                   <CEP>" & m_emit.m_enderEmit.m_CEP & "</CEP>")
    txt_xml.Append Trim("                   <cPais>" & m_emit.m_enderEmit.m_cPais & "</cPais>")
    txt_xml.Append Trim("                   <xPais>" & m_emit.m_enderEmit.m_xPais & "</xPais>")
    txt_xml.Append Trim("                   <fone>" & m_emit.m_enderEmit.m_fone & "</fone>")
    txt_xml.Append Trim("               </enderEmit>")
    txt_xml.Append Trim("               <IE>" & m_emit.m_IE & "</IE>")
    txt_xml.Append Trim("               <IM>" & m_emit.m_IM & "</IM>")    '299252
    txt_xml.Append Trim("               <CNAE>" & m_emit.m_CNAE & "</CNAE>")    '4651601
    txt_xml.Append Trim("               <CRT>" & m_emit.m_CRT & "</CRT>")    '3
    txt_xml.Append Trim("           </emit>")
    
    If Len(m_dest.m_CNPJ) > 0 Or Len(m_dest.m_CPF) > 0 Then
        txt_xml.Append Trim("           <dest>")
        If Not m_dest.m_CNPJ = "" Then
            txt_xml.Append Trim("               <CNPJ>" & m_dest.m_CNPJ & "</CNPJ>")
            txt_xml.Append Trim("               <IE>" & m_dest.m_IE & "</IE>")
        ElseIf Not m_dest.m_CPF = "" Then
            txt_xml.Append Trim("               <CPF>" & m_dest.m_CPF & "</CPF>")
            'txt_xml.Append Trim("               <IE />")
        End If
        If Len(m_dest.m_xNome) > 0 Then
            txt_xml.Append Trim("               <xNome>" & m_dest.m_xNome & "</xNome>")
        End If
        If Len(m_dest.m_enderDest.m_xLgr) > 0 Then
            txt_xml.Append Trim("               <enderDest>")
            txt_xml.Append Trim("                   <xLgr>" & m_dest.m_enderDest.m_xLgr & "</xLgr>")
            txt_xml.Append Trim("                   <nro>" & m_dest.m_enderDest.m_nro & "</nro>")
            txt_xml.Append Trim("                   <xBairro>" & m_dest.m_enderDest.m_xBairro & "</xBairro>")
            txt_xml.Append Trim("                   <cMun>" & m_dest.m_enderDest.m_cMun & "</cMun>")
            txt_xml.Append Trim("                   <xMun>" & m_dest.m_enderDest.m_xMun & "</xMun>")
            txt_xml.Append Trim("                   <UF>" & m_dest.m_enderDest.m_UF & "</UF>")
            txt_xml.Append Trim("                   <CEP>" & m_dest.m_enderDest.m_CEP & "</CEP>")
            txt_xml.Append Trim("                   <cPais>" & m_dest.m_enderDest.m_cPais & "</cPais>")
            txt_xml.Append Trim("                   <xPais>" & m_dest.m_enderDest.m_xPais & "</xPais>")
            txt_xml.Append Trim("                   <fone>" & m_dest.m_enderDest.m_fone & "</fone>")
            txt_xml.Append Trim("               </enderDest>")
            txt_xml.Append Trim("               <indIEDest>" & m_dest.m_indIEDest & "</indIEDest>")
            txt_xml.Append Trim("               <email>" & m_dest.m_email & "</email>")
        End If
        txt_xml.Append Trim("           </dest>")
    End If
    
    'NOTA FISCAL EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL
    If m_Ambiente = 12 Then
        m_cItens = m_cItens + 1
        ReDim Preserve m_Itens(m_cItens)
        For cc = m_cItens To 1 Step -1
            Set m_Itens(cc) = m_Itens(cc - 1)
        Next
        With m_Itens(0)
            .prod.m_cProd = "1"
            .prod.m_xProd = "NOTA FISCAL EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
            .prod.m_uCom = 1
            .prod.m_qCom = 1
            .prod.m_indTot = 1
            .prod.m_vProd = 1
            .prod.m_vUnCom = 1
            .prod.m_vUnTrib = 1
            .imposto.m_vTotTrib = 1
        End With
    End If

    For cc = 0 To m_cItens
        txt_xml.Append Trim("           <det nItem=""" & CStr(cc + 1) & """>")
        txt_xml.Append Trim("               <prod>")

        With m_Itens(cc).prod
            txt_xml.Append Trim("                   <cProd>" & .m_cProd & "</cProd>")
            txt_xml.Append Trim("                   <cEAN>SEM GTIN</cEAN>")
            txt_xml.Append Trim("                   <xProd>" & (.m_xProd) & "</xProd>")
            txt_xml.Append Trim("                   <NCM>" & .m_NCM & "</NCM>")
            If Len(.m_CEST) > 0 Then
                txt_xml.Append Trim("                   <CEST>" & .m_CEST & "</CEST>")
            End If
            txt_xml.Append Trim("                   <CFOP>" & .m_CFOP & "</CFOP>")
            txt_xml.Append Trim("                   <uCom>" & .m_uCom & "</uCom>")
            txt_xml.Append Trim("                   <qCom>" & .m_qCom & "</qCom>")
            txt_xml.Append Trim("                   <vUnCom>" & cFormatMoeda_NF_1(.m_vUnCom) & "</vUnCom>")
            txt_xml.Append Trim("                   <vProd>" & cFormatMoeda_NF_1(.m_vProd) & "</vProd>")
            txt_xml.Append Trim("                   <cEANTrib>SEM GTIN</cEANTrib>")
            txt_xml.Append Trim("                   <uTrib>" & .m_uTrib & "</uTrib>")
            txt_xml.Append Trim("                   <qTrib>" & .m_qTrib & "</qTrib>")
            txt_xml.Append Trim("                   <vUnTrib>" & cFormatMoeda_NF_1(.m_vUnTrib) & "</vUnTrib>")

            If .m_vFrete > 0 Then
                txt_xml.Append Trim("                   <vFrete>" & cFormatMoeda_NF_1(.m_vFrete) & "</vFrete>")
            End If
            txt_xml.Append Trim("                   <indTot>" & .m_indTot & "</indTot>")
        End With

        txt_xml.Append Trim("               </prod>")

        With m_Itens(cc).imposto
            txt_xml.Append Trim("               <imposto>")
            txt_xml.Append Trim("                   <vTotTrib>" & cFormatMoeda_NF_1(.m_vTotTrib) & "</vTotTrib>")
            
            If m_emit.m_CRT = 1 Then    'Simples Nacional
                txt_xml.Append Trim("                   <ICMS>")

                If .m_ICMS_ICMS00_CST = "500" Then
                    txt_xml.Append Trim("                       <ICMSSN500>")
                    txt_xml.Append Trim("                           <orig>0</orig>")
                    txt_xml.Append Trim("                           <CSOSN>500</CSOSN>")
                    txt_xml.Append Trim("                           <vBCSTRet>" & cFormatMoeda_NF_1(.m_ICMS_ICMS500_vBCST) & "</vBCSTRet>")
                    txt_xml.Append Trim("                           <pST>" & cFormatMoeda_NF_1(.m_pST) & "</pST>")
                    txt_xml.Append Trim("                           <vICMSSTRet>" & cFormatMoeda_NF_1(.m_vICMSSTRet) & "</vICMSSTRet>")
                    'txt_xml.Append Trim("                           <vBCFCPSTRet>0.00</vBCFCPSTRet>")
                    'txt_xml.Append Trim("                           <pFCPSTRet>0.00</pFCPSTRet>")
                    'txt_xml.Append Trim("                           <vFCPSTRet>0.00</vFCPSTRet>")
                    'txt_xml.Append Trim("                           <pRedBCEfet>0.00</pRedBCEfet>")
                    'txt_xml.Append Trim("                           <vBCEfet>0.00</vBCEfet>")
                    'txt_xml.Append Trim("                           <pICMSEfet>0.00</pICMSEfet>")
                    'txt_xml.Append Trim("                           <vICMSEfet>0.00</vICMSEfet>")
                    txt_xml.Append Trim("                       </ICMSSN500>")
                Else
                    txt_xml.Append Trim("                       <ICMSSN102>")
                    txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                    txt_xml.Append Trim("                           <CSOSN>" & .m_ICMS_ICMS00_CST & "</CSOSN>")   '
                    txt_xml.Append Trim("                       </ICMSSN102>")
                End If
                        
                txt_xml.Append Trim("                   </ICMS>")
                txt_xml.Append Trim("                   <PIS>")
                txt_xml.Append Trim("                       <PISNT>")
                txt_xml.Append Trim("                           <CST>" & .m_PIS_PISAliq_CST & "</CST>")
                txt_xml.Append Trim("                       </PISNT>")
                txt_xml.Append Trim("                   </PIS>")
                txt_xml.Append Trim("                   <COFINS>")
                txt_xml.Append Trim("                       <COFINSNT>")
                txt_xml.Append Trim("                           <CST>" & .m_COFINS_COFINSAliq_CST & "</CST>")
                txt_xml.Append Trim("                       </COFINSNT>")
                txt_xml.Append Trim("                   </COFINS>")
            
            Else
                txt_xml.Append Trim("                   <ICMS>")
                txt_xml.Append Trim("                       <ICMS00>")
                txt_xml.Append Trim("                           <orig>" & .m_ICMS_ICMS00_orig & "</orig>")
                txt_xml.Append Trim("                           <CST>" & .m_ICMS_ICMS00_CST & "</CST>")   '
                txt_xml.Append Trim("                           <modBC>" & .m_ICMS_ICMS00_modBC & "</modBC>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <pICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_pICMS) & "</pICMS>")
                txt_xml.Append Trim("                           <vICMS>" & cFormatMoeda_NF_1(.m_ICMS_ICMS00_vICMS) & "</vICMS>")
                txt_xml.Append Trim("                       </ICMS00>")
                txt_xml.Append Trim("                   </ICMS>")
                txt_xml.Append Trim("                   <IPI>")
                txt_xml.Append Trim("                       <cEnq>" & .m_IPI_cEnq & "</cEnq>")
                txt_xml.Append Trim("                       <IPINT>")
                txt_xml.Append Trim("                           <CST>" & .m_IPI_IPINT_CST & "</CST>")
                txt_xml.Append Trim("                       </IPINT>")
                txt_xml.Append Trim("                   </IPI>")
                txt_xml.Append Trim("                   <PIS>")
                txt_xml.Append Trim("                       <PISAliq>")
                txt_xml.Append Trim("                           <CST>" & .m_PIS_PISAliq_CST & "</CST>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_PIS_PISAliq_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <pPIS>" & cFormatMoeda_NF_1(.m_PIS_PISAliq_pPIS) & "</pPIS>")
                txt_xml.Append Trim("                           <vPIS>" & cFormatMoeda_NF_1(.m_PIS_PISAliq_vPIS) & "</vPIS>")
                txt_xml.Append Trim("                       </PISAliq>")
                txt_xml.Append Trim("                   </PIS>")
                txt_xml.Append Trim("                   <COFINS>")
                txt_xml.Append Trim("                       <COFINSAliq>")
                txt_xml.Append Trim("                           <CST>" & .m_COFINS_COFINSAliq_CST & "</CST>")
                txt_xml.Append Trim("                           <vBC>" & cFormatMoeda_NF_1(.m_COFINS_COFINSAliq_vBC) & "</vBC>")
                txt_xml.Append Trim("                           <pCOFINS>" & cFormatMoeda_NF_1(.m_COFINS_COFINSAliq_pCOFINS) & "</pCOFINS>")
                txt_xml.Append Trim("                           <vCOFINS>" & cFormatMoeda_NF_1(.m_COFINS_COFINSAliq_vCOFINS) & "</vCOFINS>")
                txt_xml.Append Trim("                       </COFINSAliq>")
                txt_xml.Append Trim("                   </COFINS>")
            End If

            'txt_xml.Append Trim("                   <ICMSUFDest>")
            'txt_xml.Append Trim("                       <vBCUFDest>97.68</vBCUFDest>")
            'txt_xml.Append Trim("                       <pFCPUFDest>0</pFCPUFDest>")
            'txt_xml.Append Trim("                       <pICMSUFDest>17.0000</pICMSUFDest>")
            'txt_xml.Append Trim("                       <pICMSInter>7.00</pICMSInter>")
            'txt_xml.Append Trim("                       <pICMSInterPart>40.0000</pICMSInterPart>")
            'txt_xml.Append Trim("                       <vFCPUFDest>0</vFCPUFDest>")
            'txt_xml.Append Trim("                       <vICMSUFDest>3.91</vICMSUFDest>")
            'txt_xml.Append Trim("                       <vICMSUFRemet>5.86</vICMSUFRemet>")
            'txt_xml.Append Trim("                   </ICMSUFDest>")
            txt_xml.Append Trim("               </imposto>")

        End With
        txt_xml.Append Trim("           </det>")
    Next

    txt_xml.Append Trim("           <total>")

    If m_ide.m_mod = "65" Then
        txt_xml.Append Trim("               <ICMSTot>")
        txt_xml.Append Trim("                   <vBC>" & cFormatMoeda_NF_1(m_total.m_vBC + 0) & "</vBC>")
        txt_xml.Append Trim("                   <vICMS>" & cFormatMoeda_NF_1(m_total.m_vICMS) & "</vICMS>")
        txt_xml.Append Trim("                   <vICMSDeson>" & cFormatMoeda_NF_1(m_total.m_vICMSDeson) & "</vICMSDeson>")
        txt_xml.Append Trim("                   <vFCPUFDest>0</vFCPUFDest>")
        txt_xml.Append Trim("                   <vICMSUFDest>0</vICMSUFDest>")
        txt_xml.Append Trim("                   <vICMSUFRemet>0</vICMSUFRemet>")
        txt_xml.Append Trim("                   <vFCP>0</vFCP>")
        
        txt_xml.Append Trim("                   <vBCST>0</vBCST>")
        txt_xml.Append Trim("                   <vST>" & cFormatMoeda_NF_1(m_total.m_vST) & "</vST>")
        txt_xml.Append Trim("                   <vFCPST>0</vFCPST>")
        txt_xml.Append Trim("                   <vFCPSTRet>0</vFCPSTRet>")
       '
        
        txt_xml.Append Trim("                   <vProd>" & cFormatMoeda_NF_1(m_total.m_vProd) & "</vProd>")
    
        txt_xml.Append Trim("                   <vFrete>0</vFrete>")
        txt_xml.Append Trim("                   <vSeg>0</vSeg>")
        txt_xml.Append Trim("                   <vDesc>" & cFormatMoeda_NF_1(m_total.m_vDesc) & "</vDesc>")
        txt_xml.Append Trim("                   <vII>" & cFormatMoeda_NF_1(m_total.m_vII) & "</vII>")
        txt_xml.Append Trim("                   <vIPI>" & cFormatMoeda_NF_1(m_total.m_vIPI) & "</vIPI>")
        txt_xml.Append Trim("                   <vIPIDevol>0</vIPIDevol>")
        txt_xml.Append Trim("                   <vPIS>" & cFormatMoeda_NF_1(m_total.m_vPIS) & "</vPIS>") '11.32
        txt_xml.Append Trim("                   <vCOFINS>" & cFormatMoeda_NF_1(m_total.m_vCOFINS) & "</vCOFINS>") '52.14
        txt_xml.Append Trim("                   <vOutro>" & cFormatMoeda_NF_1(m_total.m_vOutro) & "</vOutro>")
        txt_xml.Append Trim("                   <vNF>" & cFormatMoeda_NF_1(m_total.m_vProd - m_total.m_vDesc - m_total.m_vICMSDeson + m_total.m_vST + m_total.m_vFrete + m_total.m_vSeg + m_total.m_vOutro + m_total.m_vII + m_total.m_vIPI) & "</vNF>") '686.10
        m_total.m_vNF = m_total.m_vProd - m_total.m_vDesc - m_total.m_vICMSDeson + m_total.m_vST + m_total.m_vFrete + m_total.m_vSeg + m_total.m_vOutro + m_total.m_vII + m_total.m_vIPI
        
        txt_xml.Append Trim("                   <vTotTrib>" & cFormatMoeda_NF_1(m_total.m_vTotTrib) & "</vTotTrib>")
        'txt_xml.Append Trim("                   <vNF>" & cFormatMoeda_NF_1(m_total.m_vNF) & "</vNF>") '686.10
        
        
        
        txt_xml.Append Trim("               </ICMSTot>")
    ElseIf m_ide.m_mod = "55" Then
        txt_xml.Append Trim("               <ICMSTot>")
        txt_xml.Append Trim("                   <vBC>" & cFormatMoeda_NF_1(m_total.m_vBC + 0) & "</vBC>")
        txt_xml.Append Trim("                   <vICMS>" & cFormatMoeda_NF_1(m_total.m_vICMS) & "</vICMS>")
        txt_xml.Append Trim("                   <vICMSDeson>" & cFormatMoeda_NF_1(m_total.m_vICMSDeson) & "</vICMSDeson>")
        txt_xml.Append Trim("                   <vFCPUFDest>0</vFCPUFDest>")
        txt_xml.Append Trim("                   <vICMSUFDest>27.44</vICMSUFDest>")
        txt_xml.Append Trim("                   <vICMSUFRemet>41.16</vICMSUFRemet>")
        txt_xml.Append Trim("                   <vBCST>0</vBCST>")
        txt_xml.Append Trim("                   <vST>" & cFormatMoeda_NF_1(m_total.m_vST) & "</vST>")
        txt_xml.Append Trim("                   <vProd>" & cFormatMoeda_NF_1(m_total.m_vProd) & "</vProd>")
    
        txt_xml.Append Trim("                   <vFrete>" & cFormatMoeda_NF_1(m_total.m_vFrete) & "</vFrete>")
        txt_xml.Append Trim("                   <vSeg>" & cFormatMoeda_NF_1(m_total.m_vSeg) & "</vSeg>")
        txt_xml.Append Trim("                   <vDesc>" & cFormatMoeda_NF_1(m_total.m_vDesc) & "</vDesc>")
        txt_xml.Append Trim("                   <vII>" & cFormatMoeda_NF_1(m_total.m_vII) & "</vII>")
        txt_xml.Append Trim("                   <vIPI>" & cFormatMoeda_NF_1(m_total.m_vIPI) & "</vIPI>")
        txt_xml.Append Trim("                   <vPIS>" & cFormatMoeda_NF_1(m_total.m_vPIS) & "</vPIS>") '11.32
        txt_xml.Append Trim("                   <vCOFINS>" & cFormatMoeda_NF_1(m_total.m_vCOFINS) & "</vCOFINS>") '52.14
        txt_xml.Append Trim("                   <vOutro>" & cFormatMoeda_NF_1(m_total.m_vOutro) & "</vOutro>")
        
        m_total.m_vNF = m_total.m_vProd - m_total.m_vDesc - m_total.m_vICMSDeson + m_total.m_vST + m_total.m_vFrete + m_total.m_vSeg + m_total.m_vOutro + m_total.m_vII + m_total.m_vIPI
        txt_xml.Append Trim("                   <vNF>" & cFormatMoeda_NF_1(m_total.m_vNF) & "</vNF>") '686.10
        txt_xml.Append Trim("                   <vTotTrib>" & cFormatMoeda_NF_1(m_total.m_vTotTrib) & "</vTotTrib>")
    
        '(+) vProd (id:W07)
        '(-) vDesc (id:W10)
        '(-) vICMSDeson (id:W04a)
        '(+) vST (id:W06)
        '(+) vFrete (id:W08)
        '(+) vSeg (id:W09)
        '(+) vOutro (id:W15)
        '(+) vII (id:W11)
        '(+) vIPI (id:W12)
        '(+) vServ (id:W18) (*3) (NT 2011/005)
    
        txt_xml.Append Trim("               </ICMSTot>")
    End If
    txt_xml.Append Trim("           </total>")
    If m_ide.m_mod = "55" Then
        txt_xml.Append Trim("           <transp>")
        txt_xml.Append Trim("               <modFrete>0</modFrete>")
        txt_xml.Append Trim("               <transporta>")
        txt_xml.Append Trim("                   <xNome>PAC</xNome>")
        txt_xml.Append Trim("                   <IE>ISENTO</IE>")
        txt_xml.Append Trim("                   <xMun>GOIANIA</xMun>")
        txt_xml.Append Trim("                   <UF>GO</UF>")
        txt_xml.Append Trim("               </transporta>")
        txt_xml.Append Trim("               <vol>")
        txt_xml.Append Trim("                   <qVol>1</qVol>")
        txt_xml.Append Trim("                   <esp>CX</esp>")
        txt_xml.Append Trim("                   <pesoL>12.000</pesoL>")
        txt_xml.Append Trim("                   <pesoB>12.000</pesoB>")
        txt_xml.Append Trim("               </vol>")
        txt_xml.Append Trim("           </transp>")
    Else
        txt_xml.Append Trim("           <transp>")
        txt_xml.Append Trim("               <modFrete>9</modFrete>")
        txt_xml.Append Trim("           </transp>")
    End If

    If m_ide.m_mod = "55" Then
        txt_xml.Append Trim("           <cobr>")
        txt_xml.Append Trim("               <dup>")
        txt_xml.Append Trim("                   <nDup>1  000008760</nDup>")
        txt_xml.Append Trim("                   <dVenc>2016-12-27</dVenc>")
        txt_xml.Append Trim("                   <vDup>686.10</vDup>")
        txt_xml.Append Trim("               </dup>")
        txt_xml.Append Trim("           </cobr>")
        MsgBox "ajustar"
    End If
    txt_xml.Append Trim("            <pag>")
    txt_xml.Append Trim("          <detPag>")
    txt_xml.Append Trim("                <tPag>01</tPag>")
    txt_xml.Append Trim("                <vPag>" & cFormatMoeda_NF_1(m_total.m_vNF) & "</vPag>")
    txt_xml.Append Trim("              </detPag>")
    txt_xml.Append Trim("              <vTroco>0.00</vTroco>")
    txt_xml.Append Trim("            </pag>")
    
   ' txt_xml.Append Trim("           </pag>")
    'txt_xml.Append Trim("           <infAdic>")
    'txt_xml.Append Trim("           </infAdic>")
    If Len(m_infAdic.infCpl) > 0 Then
        txt_xml.Append Trim("<infAdic>")
        'txt_xml.Append Trim("   <infAdFisco>CTR  16- 0027988 ??   -    1 Item</infAdFisco>
        txt_xml.Append Trim("    <infCpl>" & m_infAdic.infCpl & "</infCpl>")
        txt_xml.Append Trim("</infAdic>")
    End If
    
    txt_xml.Append Trim("       </infNFe>")
    
    
    txt_xml.Append Trim("       <Signature xmlns=""http://www.w3.org/2000/09/xmldsig#"">")
    txt_xml.Append Trim("           <SignedInfo xmlns=""http://www.w3.org/2000/09/xmldsig#"">")
    txt_xml.Append Trim("               <CanonicalizationMethod Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>")
    txt_xml.Append Trim("               <SignatureMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#rsa-sha1""/>")
    txt_xml.Append Trim("               <Reference URI=""#" & "NFe" & m_Chave & """>")
    txt_xml.Append Trim("                   <Transforms>")
    txt_xml.Append Trim("                       <Transform Algorithm=""http://www.w3.org/2000/09/xmldsig#enveloped-signature""/>")
    txt_xml.Append Trim("                       <Transform Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>")
    txt_xml.Append Trim("                   </Transforms>")
    txt_xml.Append Trim("                   <DigestMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#sha1""/>")
    txt_xml.Append Trim("                   <DigestValue>yzGYhUx1/XYYzksWB+fPR3Qc50c=</DigestValue>")
    txt_xml.Append Trim("               </Reference>")
    txt_xml.Append Trim("           </SignedInfo>")
    txt_xml.Append Trim("           <SignatureValue></SignatureValue>")
    txt_xml.Append Trim("           <KeyInfo>")
    txt_xml.Append Trim("               <X509Data>")
    txt_xml.Append Trim("                   <X509Certificate></X509Certificate>")
    txt_xml.Append Trim("               </X509Data>")
    txt_xml.Append Trim("           </KeyInfo>")
    txt_xml.Append Trim("       </Signature>")
    txt_xml.Append Trim("   </NFe>")
    'txt_xml.Append trim( "   <protNFe xmlns=""http://www.portalfiscal.inf.br/nfe"" versao=""4.00"">"                                      )
    'txt_xml.Append trim( "       <infProt>"                                                                                               )
    'txt_xml.Append trim( "           <tpAmb>1</tpAmb>"                                                                                    )
    'txt_xml.Append trim( "           <verAplic>SVRS201611281547</verAplic>"                                                               )
    'txt_xml.Append trim( "           <chNFe>42161105865176000472550010000087601003965333</chNFe>"                                         )
    'txt_xml.Append trim( "           <dhRecbto>2016-11-29T15:47:09-02:00</dhRecbto>"                                                      )
    'txt_xml.Append trim( "           <nProt>342160141301537</nProt>"                                                                      )
    'txt_xml.Append trim( "           <digVal>N29ouIRE2DePlcXWcR3S7SSqF94=</digVal>"                                                       )
    'txt_xml.Append trim( "           <cStat>100</cStat>"                                                                                  )
    'txt_xml.Append trim( "           <xMotivo>Autorizado o uso da NF-e</xMotivo>"                                                         )
    'txt_xml.Append trim( "       </infProt>"                                                                                              )
    'txt_xml.Append trim( "   </protNFe>"                                                                                                  )
    'txt_xml.Append trim( "</nfeProc>"                                                                                                     )
    GetNota = (txt_xml)
End Function

Public Function InsertQRCode(txt_xml As String) As Boolean
Dim xml As New DOMDocument50
Dim digVal As String
Dim ooo As IXMLDOMNode
Dim txt As String
Dim txt_out As String
Dim oooSim As IXMLDOMNode
Dim Pos As Long
    If m_ide.m_mod = "55" Then
        InsertQRCode = True
    ElseIf m_ide.m_mod = "65" Then
        If xml.LoadXml(txt_xml) Then
            If ExistSelect(xml.documentElement.childNodes, oooSim, "Signature") Then
                If ExistSelect(oooSim.childNodes, ooo, "SignedInfo") Then
                    If ExistSelect(ooo.childNodes, ooo, "Reference") Then
                        If ExistSelect(ooo.childNodes, ooo, "DigestValue") Then
                            digVal = ooo.Text
                        End If
                        'retNfeStatusServico.cStat = ooo2.text
                    End If
                End If
            End If
            If Len(digVal) > 0 Then
                txt = GetQRCode(digVal)
                Pos = InStr(txt_xml, "</infNFe>")
                If Pos > 0 Then
                    txt_out = Left(txt_xml, Pos + 8)
                    txt_out = txt_out & txt
                    txt_out = txt_out & Mid(txt_xml, Pos + 9)
                    txt_xml = txt_out
                    
                    'Debug.Print Left(txt_xml, pos + 8)
                    'Debug.Print mid(txt_xml, pos + 9)
                    'Debug.Print txt
                    'Debug.Print txt_xml
                    'Debug.Print txt_out
                    'ooon.Load txt
                    'MsgBox ooon.documentElement.childNodes(1).text
                    'xml.insertBefore ooon.selectSingleNode("infNFeSupl"), xml.childNodes.Item(1)
                    'txt = Clipboard.GetText
                    'Debug.Print txt
                    InsertQRCode = True
                End If
            End If
        End If
    End If
End Function

Private Function GetQRCode(ByVal digVal As String) As String
Dim txt_xml As String
Dim txt_xml2 As String
Dim hash As New hSHA1
Dim m_HashQRCode  As String


    txt_xml = txt_xml & "<infNFeSupl>"
    txt_xml = txt_xml & "<qrCode>"
    txt_xml = txt_xml & "<![CDATA["
    
'    txt_xml2 = txt_xml2 & "?chNFe=" & m_Chave_NF
'    txt_xml2 = txt_xml2 & "&nVersao=200"
'    txt_xml2 = txt_xml2 & "&tpAmb=" & m_Ambiente
'    If Not m_dest.m_CNPJ & m_dest.m_CPF = "" Then
'        txt_xml2 = txt_xml2 & "&cDest=" & m_dest.m_CNPJ & m_dest.m_CPF
'    End If
'    txt_xml2 = txt_xml2 & "&dhEmi=" & (Base16.Encode(cFormatDate_NF_1(m_ide.m_dhEmi)))
'    txt_xml2 = txt_xml2 & "&vNF=" & cFormatMoeda_NF_1(m_total.m_vNF)
'    txt_xml2 = txt_xml2 & "&vICMS=" & cFormatMoeda_NF_1(m_total.m_vICMS)
'    txt_xml2 = txt_xml2 & "&digVal=" & Base16.Encode(digVal)
'    txt_xml2 = txt_xml2 & "&cIdToken=1"
    
    
    txt_xml2 = ""
    txt_xml2 = txt_xml2 & m_Chave_NF
    txt_xml2 = txt_xml2 & "|2"
    txt_xml2 = txt_xml2 & "|" & m_Ambiente
    txt_xml2 = txt_xml2 & "|1"
    txt_xml2 = txt_xml2 & m_CSC
    
    m_HashQRCode = hash.SHA1AA1Hash(txt_xml2)
    
    
    'txt_xml = txt_xml & "http://webas.sefaz.pi.gov.br/nfceweb-homologacao/consultarNFCe.jsf"
    'If m_UF = "PI" Then
    '    txt_xml = txt_xml & "http://www.sefaz.pi.gov.br/nfce/qrcode"
    'ElseIf m_UF = "MA" Then
    '    txt_xml = txt_xml & "http://nfce.sefaz.ma.gov.br/portal/consultarNFCe.jsp"
    'End If
    txt_xml = txt_xml & m_Endereco_QRCode
    
    'txt_xml = txt_xml & "http://www.sefaz.pi.gov.br/nfce/consulta"
    txt_xml = txt_xml & "?p="
    txt_xml = txt_xml & m_Chave_NF          'chave_acesso
    txt_xml = txt_xml & "|2"                'versao_qrcode
    txt_xml = txt_xml & "|" & m_Ambiente    'tipo_ambiente
    txt_xml = txt_xml & "|1"           'identificador_csc
    txt_xml = txt_xml & "|" & m_HashQRCode  'codigo_hash
    
    txt_xml = txt_xml & "]]>"
    txt_xml = txt_xml & "</qrCode>"
    Debug.Print txt_xml
    
'    If False Then
'
'        txt_xml = txt_xml & "<![CDATA[" & m_Endereco_QRCode
'        txt_xml2 = txt_xml2 & "?chNFe=" & m_Chave_NF
'        txt_xml2 = txt_xml2 & "&nVersao=200"
'        txt_xml2 = txt_xml2 & "&tpAmb=" & m_Ambiente
'        If m_dest.m_CNPJ & m_dest.m_CPF = "" Then
'            'txt_xml2 = txt_xml2 & "&cDest=00000000000"
'        Else
'            txt_xml2 = txt_xml2 & "&cDest=" & m_dest.m_CNPJ & m_dest.m_CPF
'        End If
'        txt_xml2 = txt_xml2 & "&dhEmi=" & (Base16.Encode(cFormatDate_NF_1(m_ide.m_dhEmi)))
'        txt_xml2 = txt_xml2 & "&vNF=" & cFormatMoeda_NF_1(m_total.m_vNF)
'        txt_xml2 = txt_xml2 & "&vICMS=" & cFormatMoeda_NF_1(m_total.m_vICMS)
'        txt_xml2 = txt_xml2 & "&digVal=" & Base16.Encode(digVal)
'        txt_xml2 = txt_xml2 & "&cIdToken=1"
'        m_HashQRCode = hash.SHA1AA1Hash(txt_xml2 & m_CSC)
'        txt_xml = txt_xml & txt_xml2
'        txt_xml = txt_xml & "&cHashQRCode=" & m_HashQRCode & "]]>"
'        txt_xml = txt_xml & "</qrCode>"
'    End If
    
    txt_xml = txt_xml & "<urlChave>"
    txt_xml = txt_xml & m_Endereco_QRCode
    txt_xml = txt_xml & "</urlChave>"
    txt_xml = txt_xml & "</infNFeSupl>"
    GetQRCode = txt_xml
End Function

Function cFormatDate_NF_1(ByVal mDate As Date) As String
    cFormatDate_NF_1 = Format(mDate, "yyyy-mm-dd") & "T" & Format(mDate, "hh:mm:ss") & "-02:00"
End Function

Function cFormatMoeda_NF_1(ByVal m_valor As Double) As String
    cFormatMoeda_NF_1 = Format(m_valor, "0.00")
    cFormatMoeda_NF_1 = Replace(cFormatMoeda_NF_1, ",", ".")
End Function

Private Sub Class_Initialize()
    m_cItens = -1
End Sub

Private Function ExistSelect(childNodes As IXMLDOMNodeList, ooo As IXMLDOMNode, ByVal mName As String) As Boolean
Dim cc As Long
    
    For cc = 0 To childNodes.Length
        If Not childNodes(cc) Is Nothing Then
            If childNodes(cc).baseName = mName Then
                Set ooo = childNodes(cc)
                ExistSelect = True
                Exit For
            End If
        End If
    Next
End Function


