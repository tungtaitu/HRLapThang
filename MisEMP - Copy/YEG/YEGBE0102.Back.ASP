<%@ Language=VBScript codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Expires = 0	
%>
<HTML>
<HEAD>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
</HEAD>
<BODY>
<%
func = request("func")
code = request("code")
index = request("index")
CurrentPage = request("CurrentPage")

tmpRec = Session("YEGBE0102")

Select Case func
	   Case "datachg"	
			tmpRec(CurrentPage,index + 1,0) = "upd"
			tmpRec(CurrentPage,index + 1,7) = code
			response.write index &"<BR>"
			response.write tmpRec(CurrentPage,index + 1,7) &"<BR>"
	
End Select

Session("YEGBE0102") = tmpRec


%>
</BODY>
</HTML>
