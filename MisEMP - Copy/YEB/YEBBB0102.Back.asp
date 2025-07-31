<%@ Language=VBScript codepage=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<%
Response.Expires = 0
session.codepage="65001"	
%>
<html>
<head>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
</head>
<body>
<%

func = request("func")
codestr01 = request("codestr01")
codestr02 = trim(request("codestr02"))
codestr03 = request("codestr03")
codestr04 = request("codestr04")

IF codestr02<>"" THEN
	codestr02 = REPLACE(codestr02, "'", "" )
	codestr02 = REPLACE (codestr02, vbCrLf ,"<br>")
	response.write "=="&codestr02& "<BR>"
END IF 

index = request("index")
CurrentPage = request("CurrentPage")
self = "admin"
'Response.Write func & "<p>"

Response.Write "CurrentPage=" & CurrentPage & "<p>"
Response.Write index & "-index <BR>"
'response.write codestr04 &"<BR>"
tmpRec = Session("YEBBB0102")

Select Case func
	   Case "del"
			tmpRec(CurrentPage,index + 1,0) = "del"
	   case "upd" 		
			tmpRec(CurrentPage,index + 1,0) = "upd" 			
			tmpRec(CurrentPage,index + 1,6) = codestr01
			tmpRec(CurrentPage,index + 1,23) = codestr02
			
End Select

response.write  tmpRec(CurrentPage,index + 1,0) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,1) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,2) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,3) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,4) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,5) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,6) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,23) &"<BR>"

Session("YEBBB0102") = tmpRec

%>
</BODY>
</HTML>

