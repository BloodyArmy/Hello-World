
Option Explicit
Dim sWebServiceURL, sContentType, sSOAPAction, sSOAPRequest
Dim oWinHttp
Dim sResponse
 
'Web Service URL
sWebServiceURL ="http://www.webserviceX.NET/stockquote.asmx"
 
'Web Service Content Type
sContentType ="text/XML"
 
'Web Service SOAP Action
sSOAPAction = "http://www.webserviceX.NET/GetQuote"
 
'Request Body
sSOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
"<soap:Body>" & _
"<GetQuote xmlns=""http://www.webserviceX.NET/"">" & _
"<symbol>MSFT</symbol>" & _
"</GetQuote>" & _
"</soap:Body>" & _
"</soap:Envelope>"
 
Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
 
'Open HTTP connection
oWinHttp.Open "POST", sWebServiceURL, False
 
'Setting request headers
oWinHttp.setRequestHeader "Content-Type", sContentType
oWinHttp.setRequestHeader "SOAPAction", sSOAPAction
 
'Send SOAP request
oWinHttp.Send sSOAPRequest
 
'Get XML Response
sResponse = oWinHttp.ResponseText
 
' Close object
Set oWinHttp = Nothing
 
'' Extract result
'Dim nPos1, nPos2
'nPos1 = InStr(sResponse, "Result>") + 7
'nPos2 = InStr(sResponse, "</")
'If nPos1 > 7 And nPos2 > 0 Then
'sResponse = Mid(sResponse, nPos1, nPos2 - nPos1)
'End If
 
' Return result
print sResponse

'
'
'Option Explicit
'Dim sWebServiceURL, sContentType, sSOAPAction, sSOAPRequest
'Dim oWinHttp
'Dim sResponse
' 
''Web Service URL
'sWebServiceURL ="http://www.w3schools.com/webservices/tempconvert.asmx"
' 
''Web Service Content Type
'sContentType ="text/XML"
' 
''Web Service SOAP Action
'sSOAPAction = "http://tempuri.org/CelsiusToFahrenheit"
' 
''Request Body
'sSOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
'"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
'"<soap:Body>" & _
'"<CelsiusToFahrenheit xmlns=""http://tempuri.org/"">" & _
'"<Celsius>25</Celsius>" & _
'"</CelsiusToFahrenheit>" & _
'"</soap:Body>" & _
'"</soap:Envelope>"
' 
'Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
' 
''Open HTTP connection
'oWinHttp.Open "POST", sWebServiceURL, False
' 
''Setting request headers
'oWinHttp.setRequestHeader "Content-Type", sContentType
'oWinHttp.setRequestHeader "SOAPAction", sSOAPAction
' 
''Send SOAP request
'oWinHttp.Send sSOAPRequest
' 
''Get XML Response
'sResponse = oWinHttp.ResponseText
' 
'' Close object
'Set oWinHttp = Nothing
' 
'' Extract result
'Dim nPos1, nPos2
'nPos1 = InStr(sResponse, "Result>") + 7
'nPos2 = InStr(sResponse, "</")
'If nPos1 > 7 And nPos2 > 0 Then
'sResponse = Mid(sResponse, nPos1, nPos2 - nPos1)
'End If
' 
'' Return result
'msgbox sResponse
'
''''Dim env
''''Dim req
'''
'''Set env = new SOAPEnvelope();
'''env.setFunctionName("GetQuote", "http://www.webserviceX.NET/");
'''env.addFunctionParameter("symbol", "IBM");
''' 
'''var req = new SOAPRequest("http://www.webservicex.net/stockquote.asmx");
'''req.setSoapAction("http://www.webserviceX.NET/GetQuote");
'''req.post(env);
'''gs.print(req.getResponseDoc()); ' print out the whole response XML
''
''Dim fso, outFile
''Set fso = CreateObject("Scripting.FileSystemObject")
''Set outFile = fso.CreateTextFile("output.txt", True)
''
''set req = CreateObject("Chilkat.HttpRequest")
''set http = CreateObject("Chilkat.Http")
''
'''  Any string unlocks the component for the 1st 30-days.
''success = http.UnlockComponent("Anything for 30-day trial")
''If (success <> 1) Then
''    MsgBox http.LastErrorText
''    WScript.Quit
''End If
''
'''  Build this XML SOAP request:
'''  <?xml version="1.0" encoding="utf-8"?>
'''  <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
'''  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
'''  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
'''    <soap:Body>
'''      <GetQuote xmlns="http://www.webserviceX.NET/">
'''        <symbol>string</symbol>
'''     </GetQuote>
'''    </soap:Body>
'''  </soap:Envelope>
''set soapReq = CreateObject("Chilkat.Xml")
''soapReq.Encoding = "utf-8"
''soapReq.Tag = "soap:Envelope"
''soapReq.AddAttribute "xmlns:xsi","http://www.w3.org/2001/XMLSchema-instance"
''soapReq.AddAttribute "xmlns:xsd","http://www.w3.org/2001/XMLSchema"
''soapReq.AddAttribute "xmlns:soap","http://schemas.xmlsoap.org/soap/envelope/"
''
''soapReq.NewChild2 "soap:Body",""
''soapReq.FirstChild2 
''soapReq.NewChild2 "GetQuote",""
''soapReq.FirstChild2 
''soapReq.AddAttribute "xmlns","http://www.webserviceX.NET/"
''soapReq.NewChild2 "symbol","MSFT"
''soapReq.GetRoot2 
''
''outFile.WriteLine(soapReq.GetXml())
''
'''  Build an SOAP request.
''req.UseXmlHttp soapReq.GetXml()
''req.Path = "/stockquote.asmx"
''
''req.AddHeader "SOAPAction","http://www.webserviceX.NET/GetQuote"
''
'''  Send the HTTP POST and get the response.  Note: This is a blocking call.
'''  The method does not return until the full HTTP response is received.
''
''domain = "www.webservicex.net"
''port = 80
''ssl = 0
''
''' resp is a Chilkat.HttpResponse
''Set resp = http.SynchronousRequest(domain,port,ssl,req)
''If (resp Is Nothing ) Then
''    outFile.WriteLine(http.LastErrorText)
''Else
''    '  The XML response is in the BodyStr property of the response object:
''    set soapResp = CreateObject("Chilkat.Xml")
''    soapResp.LoadXml resp.BodyStr
''
''    '  The response will look like this:
''    '  <?xml version="1.0" encoding="utf-8"?>
''    '  <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
''    '  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
''    '  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
''    '    <soap:Body>
''    '      <GetQuoteResponse xmlns="http://www.webserviceX.NET/">
''    '        <GetQuoteResult>string</GetQuoteResult>
''    '      </GetQuoteResponse>
''    '    </soap:Body>
''    '  </soap:Envelope>
''    '  Navigate to soap:Body
''    soapResp.FirstChild2 
''    '  Navigate to GetQuoteResponse
''    soapResp.FirstChild2 
''    '  Navigate to GetQuoteResult
''    soapResp.FirstChild2 
''
''    '  The actual XML response is the data within GetQuoteResult:
''    set xmlResp = CreateObject("Chilkat.Xml")
''    xmlResp.LoadXml soapResp.Content
''
''    '  Display the XML response:
''    outFile.WriteLine(xmlResp.GetXml())
''
''End If
''
''outFile.Close
''
''
'

