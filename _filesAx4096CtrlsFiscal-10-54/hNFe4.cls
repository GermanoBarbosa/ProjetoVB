VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public retNfeStatusServico  As New hNFe4_StatusServico
Public retNFe_EnviaLote2    As New hNFe4_EnviaLote2


Dim WithEvents e_xmldom As WinHttp.WinHttpRequest
Attribute e_xmldom.VB_VarHelpID = -1
'Dim WithEvents a as hFireControl

Dim pURL_NfeStatusServico As String, pURL_NfeStatusServico_xsd As String
Dim pURL_NFeAutorizacao_nfe As String, pURL_NFeAutorizacao_xsd As String
Dim pURL_NFeAutorizacao_nfce As String


Private mcert As String
Public NFe  As New hNFe4_Doc

Public m_Ambiente As Long
Public m_UF As String
Public m_UF_Cod As String

Public m_LastError As String


Sub SetCert(ByVal mCertificadoSerial As String)
  '  sCert: array [Boolean] of string = ('s',  '');
'  sDisp: array [Boolean] of string = ('is', 'l');
Dim txt As String
Dim m_pos As Long

    Dim i As Long  ',: Cardinal;
    Dim oNode As IXMLDOMNode
    Dim SetT As New Settings, Certs, StoreSrc As New Store, StoreDst As New Store, Cert As Certificate   ': OleVariant;
    Dim oRps As IXMLDOMNodeList, oLote As IXMLDOMNodeList, oSigs As IXMLDOMNodeList
    Dim s1 As String, s2 As String
 
    'Sett = CoSettings.Create
    SetT.EnablePromptForCertificateUI = True
    'StoreSrc = CoStore.Create
    Call StoreSrc.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_EXISTING_ONLY)
    'StoreDst = CoStore.Create
    Call StoreDst.Open(CAPICOM_CURRENT_USER_STORE, "TMP2", CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED)
    
    For i = StoreDst.Certificates.Count To 1 Step -1
        StoreDst.Remove StoreDst.Certificates(i)
    Next
    If StoreDst.Certificates.Count > 0 Then
        MsgBox "erro"
    End If
    
    'Set StoreDst = New Store
    Set Certs = StoreSrc.Certificates

    '//Remove certificados sem a private key.
    If Certs.Count > 0 Then
        Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_EXTENDED_PROPERTY, CAPICOM_PROPID_KEY_PROV_INFO)
    End If
    '//Somente certificados com data válida.
    If Certs.Count > 0 Then
        Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_TIME_VALID)
    End If
    'MsgBox Certs.item(1).SubjectName
    'Certs.Select
    
    
    If Certs.Count = 0 Then
        'MsgBox "Sem certificados"
    Else
        Set Cert = Certs.Item(1)
        For i = Certs.Count To 1 Step -1
            If Certs.Item(i).SerialNumber = mCertificadoSerial Then
                Set Cert = Certs.Item(i)
                
                
                'mcert = "LOCAL_MACHINE\My\" & Certs.Item(i).SubjectName
                StoreDst.Certificates.Clear
                StoreDst.Add Cert 'Certs.item(1)
                'mcert = "CURRENT_USER\\Personal\\" & Certs.Item(i).SubjectName
                
                txt = Certs.Item(i).SubjectName
                m_pos = InStr(txt, "CN=")
                If m_pos > 0 Then
                    txt = Mid$(txt, m_pos + 3)
                    m_pos = InStr(txt, ",")
                    If m_pos > 0 Then
                       txt = Left$(txt, m_pos - 1)
                       mcert = "CURRENT_USER\My\" & txt
                    End If
                End If
            
            End If
            'Debug.Print Certs.Item(i).PrivateKey
            'Debug.Print Certs.Item(i).PublicKey.
            'Debug.Print Certs.Item(i).SubjectName
            'Debug.Print Certs.Item(i).HasPrivateKey

        Next
        
    End If
'    'MsgBox Cert.ValidToDate
'   ' MsgBox Cert.PrivateKey
'    'Cert.HasPrivateKey
'    '//Filtra
'    'case Certs.Count of
'    '    0: Raise Exception.Create('Erro: Nenhum certificado vãlido pode ser encontrado.')
'    '    1: Cert = IInterface(Certs.Item[1]) as ICertificate2
'    '  Else: begin
'    'Try
'    '  Cert = Null
'    's1 = sCert(Certs.Count = 0)
'    's2 = sDisp(Certs.Count = 0)
'    'Certs = Certs.Select(Format('Certificado%s disponíve%s', [s1, s2]),
'    'Selecione o Certificado Digital para uso', FALSE)
'    'Cert = IInterface(Certs.Item[1]) as ICertificate2
'    ' except
'    '   on E: EOleException do
'    '     begin
'    '       i = e.errorCode
'    '       //Usuário abortou a seleção
'    '       If i = CAPICOM_E_CANCELLED Then
'    '         Abort()
'    '       Else
'    '         Raise
'    '     End
'    ' End
'    'End
'    'End
'
'    'If Not IsNull(Cert) Then
'    If Not Cert Is Nothing Then
'        'begin
'        '//Limpa do Store os Outros Certificados deixando só o selecionado
'        '//OBS: Importante só haver 1 certificado no Store Destino
'        '//     senão o XML ficará com vários X509DATA
'        StoreDst.certificates.Clear
'        StoreDst.Add Cert 'Certs.item(1)
'        'MsgBox Cert.SubjectName
'        '// Exclui todas nodes <Signature>.
'        'RemoveSignatures Doc
'
'        Set FXMLDoc = doc
'
'        '// Encontra o node <LoteRps>.
'        'MsgBox FXMLDoc.xml
'        Set oLote = FXMLDoc.selectNodes(".//Lote")
'        'MsgBox FXMLDoc.xml
'        If (oLote.Length = 0) Then
'            MsgBox "Nenhuma Tag Lote foi encontrada no XML.')"
'            Exit Sub
'        End If
'        '//Adiciona Tag de Assinatura para Lote
'
'        'Call AddDSig(Doc, oLote.item(0))
'
'        '// Encontra os nodes <InfRps>.
'        Set oRps = doc.selectNodes(".//RPS")
'        If (oRps.Length = 0) Then
'            MsgBox "Nenhuma Tag InfRps foi encontrada no XML.')"
'            Exit Sub
'        End If
'        '//Adiciona Tag de Assinatura para CADA RPS
'        If False Then
'            For i = 0 To oRps.Length - 1
'                Call AddDSig(doc, oRps(i))
'            Next
'            '//Assina cada Tag de Assinatura
'        End If
'        Call doc.setProperty("SelectionNamespaces", DS)
'        Set oSigs = doc.selectNodes(".//ds:Signature")
'
'        If oSigs.Length > 0 Then
'            For i = (oSigs.Length - 1) To 0 Step -1
'                '  begin
'                Set oNode = oSigs.Item(i)
'                Call AddSign(doc, oNode, Cert, StoreDst)
'            Next
'            'End If
'            '//Assina Tag de Assinatura do LoteRps
'            '//oSigs = Doc.selectNodes('.//ds:Signature')
'            '//if oSigs.length > 0 then
'            '//  AddSign(Doc, oSigs.item[pred(oSigs.length)], Cert, Store)
'        End If
'
'        '//if (pKeyOut = nil) then
'        '//  Exception.Create('Falha durante assinatura.\n')
'        'except
'        ' Sett = Null
'        ' StoreDst = Null
'        ' StoreSrc = Null
'        ' Raise
'        'End
'        '//Fechar Store
'        'Sett = Null
'        'StoreDst = Null
'        'StoreSrc = Null
'        'Result = True
'    End If
'    'MsgBox Right(Doc.xml, 1000)
    
End Sub

Function Inicia() As Boolean
    'mcert = "LOCAL_MACHINE\\Personal\\CN=AUTOPECAS PAX UNIAO LTDA EPP:12025605000116, OU=AR SPC PLUS, OU=RFB e-CNPJ A1, OU=Secretaria da Receita Federal do Brasil - RFB, O=ICP-Brasil, L=TERESINA, S=PI, C=BR"
    'mcert = "LOCAL_MACHINE\\My\\CN=AUTOPECAS PAX UNIAO LTDA EPP:12025605000116, OU=AR SPC PLUS, OU=RFB e-CNPJ A1, OU=Secretaria da Receita Federal do Brasil - RFB, O=ICP-Brasil, L=TERESINA, S=PI, C=BR"
    'mcert = "CURRENT_USER\Personal\AUTOPECAS PAX UNIAO LTDA EPP:12025605000116"
   ' mcert = "CURRENT_USER\My\AUTOPECAS PAX UNIAO LTDA EPP:12025605000116"
    'mcert = "CURRENT_USER\My\CN=AUTOPECAS PAX UNIAO LTDA EPP:12025605000116, OU=AR SPC PLUS, OU=RFB e-CNPJ A1, OU=Secretaria da Receita Federal do Brasil - RFB, O=ICP-Brasil, L=TERESINA, S=PI, C=BR"
    'mcert = "CURRENT_USER\My\AUTOPECAS PAX UNIAO LTDA EPP:12025605000116"
    ' mcert = "LOCAL_MACHINE\Personal\AUTOPECAS PAX UNIAO LTDA"
    'mcert = "LOCAL_MACHINE\\My\\AUTOPECAS PAX UNIAO LTDA EPP:12025605000116"
    'Class_Initialize
    'List1.AddItem NfeStatusServico
    'List1.AddItem EnviaLote2
    
    'OTICA PAX UNIAO LTDA - EPP:12020938000152
    'mcert = "CURRENT_USER\My\CAFEME COMERCIO DE TECIDOS LTDA:16655228000150"
    'mcert = "CURRENT_USER\My\AUTOPECAS PAX UNIAO LTDA EPP:12025605000116"
    Inicia = True
    'Debug.Print mcert
    
    'http://hom.nfe.fazenda.gov.br/PORTAL/WebServices.aspx#SVC-RS
    
    If m_Ambiente = 1 Then
        pURL_NfeStatusServico = "https://nfce.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
        ''pURL_NFeAutorizacao_nfe = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao.asmx"
        pURL_NFeAutorizacao_nfce = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
    Else
        pURL_NfeStatusServico = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
        pURL_NfeStatusServico = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
        'pURL_NFeAutorizacao_nfe = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao.asmx"
        pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
    End If
    
    
 
End Function

Public Sub Initialize(ByVal m_path_Schema As String, ByVal m_path_app As String)
Dim m_Caminho As String
    'pURL_NfeStatusServico = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico2.asmx"
    'pURL_NFeAutorizacao_nfe = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao.asmx"
    'pURL_NFeAutorizacao_nfce = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao.asmx"
    
    
    'RecepcaoEvento  1.00    https://nfe-homologacao.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento.asmx
    'NfeConsultaCadastro 2.00    https://cad.svrs.rs.gov.br/ws/cadconsultacadastro/cadconsultacadastro2.asmx
    'NfeInutilizacao 4.00    https://nfe-homologacao.svrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao2.asmx
    'NfeConsultaProtocolo    4.00    https://nfe-homologacao.svrs.rs.gov.br/ws/NfeConsulta/NfeConsulta2.asmx
    '*NfeStatusServico    4.00    https://nfe-homologacao.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico2.asmx
    'NFeAutorizacao  4.00    https://nfe-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao.asmx
    'NFeRetAutorizacao   4.00    https://nfe-homologacao.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao.asmx
    If IsIDE Then
        m_Caminho = "G:\sistema\os\NFe\Schema-4.00\"
        m_path = "G:\sistema\pra\"
    Else
        m_Caminho = App.Path & "..\os\NFe\Schema-4.00\"
        m_path = App.Path & "\"
    End If
    
    m_Caminho = m_path_Schema
    m_path = m_path_app
    
    pURL_NFeAutorizacao_xsd = m_Caminho & "enviNFe_v4.00.xsd"
    pURL_NfeStatusServico_xsd = m_Caminho & "consStatServ_v4.00.xsd"
    
    If Not FileExists(pURL_NFeAutorizacao_xsd) Then
        MsgBox "Schema 4.00 NÃO encontrado em: " & m_Caminho
    End If
    'If Not FileExists(pURL_NFeAutorizacao_xsd) Then
    '    MsgBox "Erro 2"
    'End If
End Sub

Function NfeStatusServico() As Boolean
          Dim xmldoc As New WinHttpRequest
          Dim txt As String
          Dim xmlnf As New MSXML2.DOMDocument50
          Dim xmldom_resp As New MSXML2.DOMDocument50
          Dim txt_xml As String

10        On Error GoTo NfeStatusServico_Error
'20        txt_xml = "<?xml version=""1.0"" encoding=""UTF-8""?>"
'30        txt_xml = txt_xml & "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NfeStatusServico2"">"
'40        txt_xml = txt_xml & "<soap:Header>"
'50        txt_xml = txt_xml & "<nfeCabecMsg>"
'60        txt_xml = txt_xml & "<cUF>"
'70        txt_xml = txt_xml & m_UF_Cod
'80        txt_xml = txt_xml & "</cUF>" '41 PIAUI
'90        txt_xml = txt_xml & "<versaoDados>4.00</versaoDados>"
'100       txt_xml = txt_xml & "</nfeCabecMsg>"
'110       txt_xml = txt_xml & "</soap:Header>"
'120       txt_xml = txt_xml & "<soap:Body>"
'130       txt_xml = txt_xml & "<nfeDadosMsg>"
'140       txt_xml = txt_xml & "<consStatServ versao=""4.00"" xmlns=""http://www.portalfiscal.inf.br/nfe""><tpAmb>"
'150       txt_xml = txt_xml & m_Ambiente
'160       txt_xml = txt_xml & "</tpAmb><cUF>"
'170       txt_xml = txt_xml & m_UF_Cod
'180       txt_xml = txt_xml & "</cUF><xServ>STATUS</xServ></consStatServ>"
'190       txt_xml = txt_xml & "</nfeDadosMsg>"
'200       txt_xml = txt_xml & "</soap:Body>"
'210       txt_xml = txt_xml & "</soap:Envelope>"
          
'220       If True Then
230       txt_xml = ""
240       txt_xml = txt_xml & "<?xml version=""1.0"" encoding=""utf-8""?>"
250       txt_xml = txt_xml & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
260       txt_xml = txt_xml & "<soap12:Body>"
270       txt_xml = txt_xml & "<nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NFeStatusServico4""> "
          
280       txt_xml = txt_xml & "<consStatServ versao=""4.00"" xmlns=""http://www.portalfiscal.inf.br/nfe""><tpAmb>"
290       txt_xml = txt_xml & m_Ambiente
300       txt_xml = txt_xml & "</tpAmb><cUF>"
310       txt_xml = txt_xml & m_UF_Cod
320       txt_xml = txt_xml & "</cUF><xServ>STATUS</xServ></consStatServ>"
          
          
330       txt_xml = txt_xml & "</nfeDadosMsg>"
340       txt_xml = txt_xml & "</soap12:Body>"
350       txt_xml = txt_xml & "</soap12:Envelope>"
'360       End If
         ' Debug.Print txt_xml
370       If xmlnf.LoadXml(txt_xml) Then
380           txt = ValidaSchema(xmlnf, "http://www.portalfiscal.inf.br/nfe", pURL_NfeStatusServico_xsd, False)
390           If txt = "Validado" Then
              
400               Call xmldoc.Open("POST", pURL_NfeStatusServico & "?wsdl", False)
410               Call xmldoc.SetRequestHeader("content-type", "text/xml")
                  'Call xmldoc.setRequestHeader("content-type", "application/soap+xml")
420               Call xmldoc.SetRequestHeader("Cache-Control", "no-cache")
                  
430               Call xmldoc.SetRequestHeader("soapAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeStatusServico4/nfeStatusServicoNF")
                  'Call xmldoc.setRequestHeader("soapAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeAutorizacao4/nfeAutorizacaoLote")
                  
                 ' Call xmldoc.setRequestHeader("SOAPAction", """")
                  
440               xmldoc.SetClientCertificate (mcert)
                 ' xmldoc.Option(WinHttpRequestOption_SelectCertificate) = (mcert)
                  
                  'https://www.experts-exchange.com/questions/28821379/winhttp-dll-not-working-with-TLS1-1-and-TLS1-2-on-windows-2008-r2-x64.html
              '    xmldoc.Option(WinHttpRequestOption_SecureProtocols) = 2048      'TLS 1.2
                  
450               xmldoc.Send (xmlnf.xml)
460               If xmldom_resp.LoadXml(xmldoc.ResponseText) Then
                      'Debug.Print "---------------------"
                      'Debug.Print pURL_NfeStatusServico
                      'Debug.Print xmldom_resp.xml
                      'Debug.Print "---------------------"

                      Dim ooo As IXMLDOMNode
                      Dim ooo2 As IXMLDOMNode
470                   If ExistSelect(xmldom_resp.documentElement.childNodes, ooo, "Body") Then
480                       If ExistSelect(ooo.childNodes, ooo, "nfeResultMsg") Then
490                           If ExistSelect(ooo.childNodes, ooo, "retConsStatServ") Then
500                               If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
510                                   retNfeStatusServico.cStat = ooo2.Text
520                               End If
530                               If ExistSelect(ooo.childNodes, ooo2, "tpAmb") Then
540                                   retNfeStatusServico.tpAmb = ooo2.Text
550                               End If
560                               If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
570                                   retNfeStatusServico.xMotivo = ooo2.Text
580                               End If
590                           End If
600                       End If
610                   End If
                  Else
                    MsgBox xmldom_resp.parseError.srcText, , "parseError - 2"
                    
620               End If
630               NfeStatusServico = True
640           Else
650               m_LastError = txt
660           End If
670       Else
680           MsgBox xmlnf.parseError.srcText, , "parseError"
690           SetFileBytes App.Path & "\error_1.txt", txt_xml
700           SetFileBytes App.Path & "\error_2.txt", xmlnf.Text
710       End If
720       Set xmldoc = Nothing
730       On Error GoTo 0
740       Exit Function
750       Resume Next
NfeStatusServico_Error:
760       Debug.Print Err.Number, Err.Description
770       If Err.Number = -2147012711 Then
780           MsgBox "Erro de certificado"
790       ElseIf Err.Number = -2147012852 Then
800           MsgBox "É necessário um certificado para concluir a autenticação do cliente"
810       ElseIf Err.Number = -2147012851 Then
820           MsgBox "A autoridade de certificação não é válida ou está incorreta"
              'https://www.projetoacbr.com.br/forum/topic/27273-a-autoridade-de-certifica%C3%A7%C3%A3o-n%C3%A3o-%C3%A9-valida-ou-esta-incorreta/
830       Else
840           MsgBox "Error linha: " & Erl & "," & Err.Number & " (" & Err.Description & ") in procedure NfeStatusServico of Módulo mMainTeste"
850           Debug.Print Err.Number,
860       End If
End Function

Function EnviaLote2(ByVal xml As String, xml_aut As String, ByVal mCertificadoSel As String) As Boolean
   On Error GoTo EnviaLote2_Error

10
          Dim xmlnf As New MSXML2.DOMDocument50
          Dim xmldoc As New WinHttp.WinHttpRequest    'MSXML2.XMLHTTP50
          Dim xml_ass As String
          Dim txt_xml As New hStringBuilder
          Dim txt As String
          Dim xmldom_resp As New MSXML2.DOMDocument50
          Dim txt_out As String
                   
20        txt_xml.Clear
          
30        txt_xml.Append "<?xml version=""1.0"" encoding=""utf-8""?>"
40        txt_xml.Append "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
50        txt_xml.Append "<soap12:Body>"
60        txt_xml.Append "<nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/NFeAutorizacao4"">"
          
70        txt_xml.Append "<enviNFe xmlns=""http://www.portalfiscal.inf.br/nfe"" versao=""4.00"">"
          
80        txt_xml.Append "<idLote>1</idLote>"
90        txt_xml.Append "<indSinc>1</indSinc>"
100       txt_xml.Append Mid(xml, 1, Len(xml) - 2)
110       txt_xml.Append "</enviNFe>"
          
120       txt_xml.Append "</nfeDadosMsg>"
130       txt_xml.Append "</soap12:Body>"
140       txt_xml.Append "</soap12:Envelope>"
          
          Dim ooo_insert As IXMLDOMNode
          
          Dim m_File As String
150       If IsIDE Then
160           m_File = "g:\sistema\pra\NFe\temp\" & Format(Time, "hh-mm-ss") & ".xml"
170           Clipboard.Clear
180           Clipboard.SetText txt_xml, vbCFText
190       Else
200           m_File = m_path & "\NFe\temp\" & Format(Time, "hh-mm-ss") & ".xml"
210       End If
220       Call SetFileBytes(m_File, txt_xml.toString)
          
          
          
          'leiauteNFe_v3 0.1
          'xmlnf.loadXML (xml)
          'txt = ValidaSchema(xmlnf, "http://www.portalfiscal.inf.br/nfe", App.Path & "\NFe\Schema-4.00\leiauteNFe_v4.00.xsd", False)
          'If txt = "Validado" Then
          'Else
          '    Debug.Print txt
          'End If
          
230       If xmlnf.LoadXml(txt_xml) Then
          'If xmlnf.Load(m_File) Then
              'MsgBox xmlnf.prefix
              'Set ooo_insert = xmlnf.nodeName
              'xmlnf.appendChild ooo_insert, xml_ass
240           txt = ValidaSchema(xmlnf, "http://www.portalfiscal.inf.br/nfe", pURL_NFeAutorizacao_xsd, False)
250           If txt = "Validado" Then
260               If InStr(xml, "<mod>55</mod>") > 0 Then
270                   Call xmldoc.Open("POST", pURL_NFeAutorizacao_nfe & "?wsdl", False)
280               ElseIf InStr(xml, "<mod>65</mod>") > 0 Then
290                   Call xmldoc.Open("POST", pURL_NFeAutorizacao_nfce & "?wsdl", False)
300               End If
                  'Call xmldoc.setRequestHeader("content-type", "application/soap+xml")
                  
310               Debug.Print pURL_NFeAutorizacao_nfce
                  
                  'Call xmldoc.Open("POST", pURL_NfeStatusServico & "?wsdl", False)
                  'Call xmldoc.setRequestHeader("content-type", "text/xml")
                  'Call xmldoc.setRequestHeader("content-type", "application/soap+xml")
                  'Call xmldoc.setRequestHeader("Cache-Control", "no-cache")
                  'Call xmldoc.setRequestHeader("SOAPAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeStatusServico4/nfeStatusServicoNF")
                 '' Call xmldoc.setRequestHeader("SOAPAction", """")
                  'xmldoc.SetClientCertificate (mcert)
                  
320               Call xmldoc.SetRequestHeader("content-type", "application/soap+xml; charset=utf-8")
                  
                  
                  'Call xmldoc.setRequestHeader("content-type", "text/xml")
                 'Call xmldoc.setRequestHeader("content-type", "application/soap+xml")
330               Call xmldoc.SetRequestHeader("Cache-Control", "no-cache")
340               Call xmldoc.SetRequestHeader("soapAction", "http://www.portalfiscal.inf.br/nfe/wsdl/NFeAutorizacao4/nfeAutorizacaoLote")
350               xmldoc.SetClientCertificate (mcert)
               '   xmldoc.Option(WinHttpRequestOption_SelectCertificate) = (mcert) 'mcert
                  
                  'https://www.experts-exchange.com/questions/28821379/winhttp-dll-not-working-with-TLS1-1-and-TLS1-2-on-windows-2008-r2-x64.html
                  'https://support.microsoft.com/pt-br/help/3140245/update-to-enable-tls-1-1-and-tls-1-2-as-default-secure-protocols-in-wi
360               '2019-01-30 xmldoc.Option(WinHttpRequestOption_SecureProtocols) = 2048      'TLS 1.2
                  'MsgBox xmldoc.Option(WinHttpRequestOption_SecureProtocols)
                   
                  'Call File.SetFileBytes(m_path & "\teste\xmlnf_" & Format(time, "hh-mm-ss") & ".xml", ClearXML(xmlnf.xml))
                  
                  'xmldoc.send ClearXML(xmlnf.xml)
370               xmldoc.Send (xmlnf.xml)
380               If xmldom_resp.LoadXml(xmldoc.ResponseText) Then
                      'Debug.Print "---------------------"
390                   Debug.Print xmldom_resp.Text
400                   Debug.Print xmldom_resp.xml
                      'Debug.Print "---------------------"
                      Dim ooo As IXMLDOMNode
                      Dim ooo2 As IXMLDOMNode
                      Dim oooTemp As IXMLDOMNode
                      Dim xmldom_resp2 As New MSXML2.DOMDocument50
410                   If Not xmldom_resp.documentElement Is Nothing Then
420                       If IsIDE Then
430                           xmldom_resp.Save "g:\sistema\pra\NFe\Resp\" & NFe.m_Chave_NF & "-aut.xml"
440                       Else
450                           xmldom_resp.Save m_path & "\NFe\Resp\" & NFe.m_Chave_NF & "-aut.xml"
460                       End If
                          'xml_aut = xmldom_resp.xml
470                       If ExistSelect(xmldom_resp.documentElement.childNodes, ooo, "Body") Then
480                           If ExistSelect(ooo.childNodes, ooo, "nfeResultMsg") Then
490                               If ExistSelect(ooo.childNodes, ooo, "retEnviNFe") Then
500                                   If ExistSelect(ooo.childNodes, ooo2, "xMotivo") Then
510                                       retNFe_EnviaLote2.xMotivo = ooo2.Text
520                                       Debug.Print ooo.Text
                                          
530                                   End If
540                                   If ExistSelect(ooo.childNodes, ooo2, "cStat") Then
550                                       retNFe_EnviaLote2.cMotivo = ooo2.Text
560                                       Debug.Print ooo2.Text
570                                   End If
580                                   If ExistSelect(ooo.childNodes, ooo2, "protNFe") Then
590                                       xml_aut = ooo2.xml
600                                       If xmldom_resp2.LoadXml(ooo2.xml) Then
610                                           If ExistSelect(xmldom_resp2.childNodes, ooo2, "protNFe") Then
                                                '  MsgBox ooo2.xml
620                                               If ExistSelect(ooo2.childNodes, ooo2, "infProt") Then
630                                                   If ExistSelect(ooo2.childNodes, oooTemp, "xMotivo") Then
640                                                       retNFe_EnviaLote2.protNFe.xMotivo = oooTemp.Text
                                                          'Set retNFe_EnviaLote2.protNFe = New hNFe4_EnviaLote2_protNFe
650                                                   End If
660                                                   If ExistSelect(ooo2.childNodes, oooTemp, "cStat") Then
670                                                       retNFe_EnviaLote2.protNFe.cStat = oooTemp.Text
680                                                   End If
690                                               End If
700                                           End If
710                                       End If
                                          
                                          
                                          'retNFe_EnviaLote2.protNFe = ooo2.text
                                          
720                                   End If
730                               End If
740                           End If
750                       End If
760                   End If
770               End If
780               EnviaLote2 = True
790           Else
800               Debug.Print txt
810               m_LastError = txt
820           End If
830       Else
840           MsgBox "Não carregou o xml"
850       End If
          'Set xmldom = Nothing
860       Set xmldoc = Nothing


   On Error GoTo 0
   Exit Function

EnviaLote2_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EnviaLote2 of Módulo de classe hNFe4 " & Erl

End Function

Function GetNota(m_xml As String, m_Chave As String, m_NF_NUM_NOTA As Long) As Boolean
    m_xml = NFe.GetNota(m_Chave, m_NF_NUM_NOTA)
    GetNota = True
End Function

Private Function ValidaSchema(ByVal docXML As DOMDocument50, ByVal strUrn As String, ByVal strXsdArquivo As String, ByVal comAssinatura As Boolean) As String
' cria um cache schema e adiciona o arquivo strXsdArquivo.
    Dim Xs As New XMLSchemaCache50
    Dim lngTam As Long, lngTam2 As Long, strCampoErrado As String
    ValidaSchema = True
    Xs.Add strUrn, strXsdArquivo

    ' cria um XML DOMDocument object.
    Dim xd As New DOMDocument50

    ' adiciona o schemaCache ao documento.
    Set xd.schemas = Xs

    ' Carrega o documento docXML.
    xd.async = False
    xd.LoadXml docXML.xml

    ' Return validation results in message to the user.
    If xd.parseError.errorCode <> 0 Then

        If Not comAssinatura Then
            'Ignorar erro de assinatura
            If InStr(1, UCase(xd.parseError.reason), "SIGNATURE") > 0 Then
                ValidaSchema = "Validado"
                Exit Function
            End If
        End If

        ValidaSchema = "Erro na validação: " & vbCrLf
        ValidaSchema = ValidaSchema & "- Código    : " & xd.parseError.errorCode & vbCrLf
        lngTam = InStr(1, xd.parseError.reason, "enumeration")
        If lngTam > 0 Then
            lngTam = InStr(1, xd.parseError.reason, "}")
            lngTam2 = InStr(lngTam + 1, xd.parseError.reason, "'")
            lngTam2 = lngTam2 - lngTam - 1
            strCampoErrado = Mid(xd.parseError.reason, lngTam + 1, lngTam2)
            ValidaSchema = ValidaSchema & "- Descrição: O Campo " & Chr(34) & strCampoErrado & Chr(34) & " é Inválido." & vbCrLf & "Por Favor, Verifique os Dados de sua Nota Fiscal." & vbCrLf

            ValidaSchema = ValidaSchema & "- Linha    : " & xd.parseError.Line & vbCrLf
            ValidaSchema = ValidaSchema & "- Posição  : " & xd.parseError.linepos
        Else
            ValidaSchema = ValidaSchema & "- Descrição: " & xd.parseError.reason & vbCrLf
            ValidaSchema = ValidaSchema & "- Fonte    : " & Chr(10) & Replace(xd.parseError.srcText, "><", "><" & vbCrLf) & vbCrLf
            ValidaSchema = ValidaSchema & "- Linha    : " & xd.parseError.Line
        End If
        'FunValidaSchema = FunValidaSchema & "- Fonte    : " & xd.parseError.srcText & vbCrLf
        'FunValidaSchema = FunValidaSchema & "- Linha    : " & xd.parseError.Line
        '      FunValidaSchema = False


    Else
        ValidaSchema = "Validado"
    End If

End Function

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

Function GetCert(ByVal mCertificadoSerial As String) As Certificate
  '  sCert: array [Boolean] of string = ('s',  '');
'  sDisp: array [Boolean] of string = ('is', 'l');
    Dim i As Long  ',: Cardinal;
    Dim oNode As IXMLDOMNode
    Dim SetT As New Settings, Certs, StoreSrc As New Store, StoreDst As New Store, Cert As Certificate   ': OleVariant;
    Dim oRps As IXMLDOMNodeList, oLote As IXMLDOMNodeList, oSigs As IXMLDOMNodeList
    Dim s1 As String, s2 As String
 
    'Sett = CoSettings.Create
    SetT.EnablePromptForCertificateUI = True
    'StoreSrc = CoStore.Create
    Call StoreSrc.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_EXISTING_ONLY)
    'StoreDst = CoStore.Create
    Call StoreDst.Open(CAPICOM_CURRENT_USER_STORE, "TMP2", CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED)
    
    For i = StoreDst.Certificates.Count To 1 Step -1
        StoreDst.Remove StoreDst.Certificates(i)
    Next
    If StoreDst.Certificates.Count > 0 Then
        MsgBox "erro"
    End If
    
    'Set StoreDst = New Store
    Set Certs = StoreSrc.Certificates

    '//Remove certificados sem a private key.
    If Certs.Count > 0 Then
        Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_EXTENDED_PROPERTY, CAPICOM_PROPID_KEY_PROV_INFO)
    End If
    '//Somente certificados com data válida.
    If Certs.Count > 0 Then
        Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_TIME_VALID)
    End If
    'MsgBox Certs.item(1).SubjectName
    'Certs.Select
    
    
    If Certs.Count = 0 Then
        'MsgBox "Sem certificados"
    Else
        Set Cert = Certs.Item(1)
        For i = Certs.Count To 1 Step -1
            If Certs.Item(i).SerialNumber = mCertificadoSerial Then
                Set GetCert = Certs.Item(i)
                
                Exit For
                'mcert = "LOCAL_MACHINE\My\" & Certs.Item(i).SubjectName
                StoreDst.Certificates.Clear
                StoreDst.Add Cert 'Certs.item(1)
                mcert = "CURRENT_USER\\TMP2\\" & Certs.Item(i).SubjectName
            End If
            Debug.Print Certs.Item(i).PrivateKey
            'Debug.Print Certs.Item(i).PublicKey.
            Debug.Print Certs.Item(i).SubjectName
            Debug.Print Certs.Item(i).HasPrivateKey
             
        Next
        
    End If
End Function

