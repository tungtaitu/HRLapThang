<%@ Language=VBScript codepage=950%>
<%
Response.Expires = 0	
%>
<html>
<head>
<meta http-equiv="refresh">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">	
</head>
<body>
<%

func = request("func")
codestr01 = request("codestr01")
codestr02 = request("codestr02")
codestr03 = request("codestr03")
codestr04 = request("codestr04")

index = request("index")
CurrentPage = request("CurrentPage")
self = "admin"
'Response.Write func & "<p>"

'Response.Write currentPage & "<p>"
Response.Write index & "-index <p>"

tmpRec = Session("ADMIN01")

Select Case func
	   Case "del"
			tmpRec(CurrentPage,index + 1,0) = "del"
	   case "upd" 		
			tmpRec(CurrentPage,index + 1,0) = "upd" 			
			tmpRec(CurrentPage,index + 1,1) = codestr02
			tmpRec(CurrentPage,index + 1,2) = codestr03
			tmpRec(CurrentPage,index + 1,3) = codestr04
			tmpRec(CurrentPage,index + 1,4) = codestr01
End Select

response.write  tmpRec(CurrentPage,index + 1,1) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,2) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,3) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,4) &"<BR>"

Session("ADMIN01") = tmpRec
%>
</BODY>
</HTML>

