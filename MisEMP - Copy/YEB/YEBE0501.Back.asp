<%@ Language=VBScript codepage=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<%
Response.Expires = 0
session.codepage="65001"	
self="YEBE0501"
%>
<html>
<head>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
</head>
<body>
<%
'Set conn = GetSQLServerConnection()
func = request("func")
codestr01 = request("code1")
codestr02 = trim(request("code2"))
codestr03 = request("code3")
codestr04 = request("code4")
codestr05 = request("code5")
codestr06 = request("code6")
codestr07 = request("code7")  

index = request("index")
CurrentPage = request("CurrentPage")

'Response.Write func & "<p>"

Response.Write "CurrentPage=" & CurrentPage & "<p>"
Response.Write index & "-index <BR>"
'response.write codestr04 &"<BR>"
tmpRec = Session("YEBE0501") 


Select Case func
		Case "del"
			tmpRec(CurrentPage,index + 1,0) = "del"
		Case "no"
			tmpRec(CurrentPage,index + 1,0) = "no"
		case "datachg" 		
			tmpRec(CurrentPage,index + 1,0) = "upd" 			
			tmpRec(CurrentPage,index + 1,3) = codestr01
			tmpRec(CurrentPage,index + 1,16) = codestr02
			tmpRec(CurrentPage,index + 1,17) = codestr03
			tmpRec(CurrentPage,index + 1,18) = codestr05
			tmpRec(CurrentPage,index + 1,19) = codestr06
			tmpRec(CurrentPage,index + 1,20) = codestr04
			tmpRec(CurrentPage,index + 1,21) = codestr07
			
		Case "T1Y"			
			tmpRec(CurrentPage,index + 1,4) = "Y"
		Case "T1N"			
			tmpRec(CurrentPage,index + 1,4) = ""			

		Case "T2Y"			
			tmpRec(CurrentPage,index + 1,5) = "Y"
		Case "T2N"			
			tmpRec(CurrentPage,index + 1,5) = ""			
			
		Case "T3Y"			
			tmpRec(CurrentPage,index + 1,6) = "Y"
		Case "T3N"			
			tmpRec(CurrentPage,index + 1,6) = ""			
			
		Case "T4Y"			
			tmpRec(CurrentPage,index + 1,7) = "Y"
		Case "T4N"			
			tmpRec(CurrentPage,index + 1,7) = ""			
			
		Case "T5Y"			
			tmpRec(CurrentPage,index + 1,8) = "Y"
		Case "T5N"			
			tmpRec(CurrentPage,index + 1,8) = ""			
			
		Case "T6Y"			
			tmpRec(CurrentPage,index + 1,9) = "Y"
		Case "T6N"			
			tmpRec(CurrentPage,index + 1,9) = ""			
			
		Case "T7Y"			
			tmpRec(CurrentPage,index + 1,10) = "Y"
		Case "T7N"			
			tmpRec(CurrentPage,index + 1,10) = ""			
			
		Case "T8Y"			
			tmpRec(CurrentPage,index + 1,11) = "Y"
		Case "T8N"			
			tmpRec(CurrentPage,index + 1,11) = ""			
			
		Case "T9Y"			
			tmpRec(CurrentPage,index + 1,12) = "Y"
		Case "T9N"			
			tmpRec(CurrentPage,index + 1,12) = ""																								

		Case "T10Y"			
			tmpRec(CurrentPage,index + 1,13) = "Y"
		Case "T10N"			
			tmpRec(CurrentPage,index + 1,13) = ""		

		Case "T11Y"			
			tmpRec(CurrentPage,index + 1,14) = "Y"
		Case "T11N"			
			tmpRec(CurrentPage,index + 1,14) = ""		

		Case "T12Y"			
			tmpRec(CurrentPage,index + 1,15) = "Y"
		Case "T12N"			
			tmpRec(CurrentPage,index + 1,15) = ""											

End Select

response.write  tmpRec(CurrentPage,index + 1,0) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,1) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,2) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,3) &"<BR>"
response.write  "4=" & tmpRec(CurrentPage,index + 1,4) &"<BR>"
response.write  "5=" & tmpRec(CurrentPage,index + 1,5) &"<BR>"
response.write  "6=" & tmpRec(CurrentPage,index + 1,6) &"<BR>"
response.write  "7=" & tmpRec(CurrentPage,index + 1,7) &"<BR>"
response.write  "8=" & tmpRec(CurrentPage,index + 1,8) &"<BR>"
response.write  "9=" & tmpRec(CurrentPage,index + 1,9) &"<BR>"
response.write  "10=" & tmpRec(CurrentPage,index + 1,10) &"<BR>"
response.write  "11=" & tmpRec(CurrentPage,index + 1,11) &"<BR>"
response.write  "12=" & tmpRec(CurrentPage,index + 1,12) &"<BR>"
response.write  "13=" & tmpRec(CurrentPage,index + 1,13) &"<BR>"
response.write  "14=" & tmpRec(CurrentPage,index + 1,14) &"<BR>"
response.write  "15=" & tmpRec(CurrentPage,index + 1,15) &"<BR>"
response.write  "18=" & tmpRec(CurrentPage,index + 1,18) &"<BR>"
response.write  "20=" & tmpRec(CurrentPage,index + 1,20) &"<BR>"


Session("YEBE0501") = tmpRec
%>
</BODY>
</HTML>

