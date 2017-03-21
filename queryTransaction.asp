<%
transaction_id = "183828b3-e573-478f-a311-e4a88d5b89c2"

RANDOMIZE
 
terminal_id = "0962922"
merchant_id=1
user="israel"
password="Israeli1."
cg_gateway_url="https://cguat2.creditguard.co.il/xpo/Relay"

poststring="user=" & user &_
			"&password=" & password &_
			"&int_in=<ashrait>" &_
							"<request>" &_
							 "<language>HEB</language>" &_
							 "<command>inquireTransactions</command>" &_
							 "<inquireTransactions>" &_
 							  "<terminalNumber>" & terminal_id & "</terminalNumber>" &_
							  "<mainTerminalNumber/>" &_
							  "<queryName>mpiTransaction</queryName>" &_
							  "<mid>" & merchant_id & "</mid>" &_
							  "<mpiTransactionId>" & transaction_id & "</mpiTransactionId>" &_
							  "<userData1/>" &_
							  "<userData2/>" &_
							  "<userData3/>" &_
							  "<userData4/>" &_
							  "<userData5/>" &_
							 "</inquireTransactions>" &_
							"</request>" &_
						   "</ashrait>"
						  
' HTTP POST Preparing and posting
Dim objWinHttp
Dim strResponseStatus
Dim strResponseText
' Create an instance of our HTTP object
Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
' Open a connection to the server
'   .Open(bstrMethod, bstrUrl [, varAsync])
objWinHttp.Open "POST", cg_gateway_url, False
' Set the content type header of our request to indicate
' the body of our request will contain form data.
'   .SetRequestHeader(bstrHeader, bstrValue)
objWinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
' Send the request to the server.  Form data is sent in
' the body of the request.  Here I'm simply sending a
' name and a date.  You should URLEncode any data that
' contains spaces or special characters.
'   .Send(varBody)
		
objWinHttp.Send poststring

' Get the server's response status
strResponseStatus = objWinHttp.Status & " " & objWinHttp.StatusText
' Get the text of the response
strResponseText = objWinHttp.ResponseText
' Dispose of our object now that we're done with it
Set objWinHttp = Nothing
			
' Return the Transaction ID from the MPI Server
' Full Example: postRequest = "Sending Post request:<br />" & finalRequest & "<br/><br/>Server Response: <br/>" & strResponseText

Dim xmlDoc
Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
xmlDoc.loadXML(strResponseText)
xmlDoc.async = False
Response.write(xmlDoc.getElementsByTagName("statusText").item(0).Text)
%>
