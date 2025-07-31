<%@ Language=VBScript codepage=65001%>
<%
Response.Expires = 0	
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<%
func = request("func")
tblcd = request("tblcd")
tbldesc = request("tbldesc")
index = request("index")
CurrentPage = request("CurrentPage")

Response.Write func & "<p>"
Response.Write tblcd & "-2" & "<p>"
Response.Write tbldesc & "-3" & "<p>"
Response.Write currentPage & "<p>"
Response.Write index & "-index <p>"

tmpRec = Session("YDBSB0001EMP")

Select Case func
	   Case "tblcd_change"
			tmpRec(CurrentPage,index + 1,0) = "update"
			tmpRec(CurrentPage,index + 1,1) = tblcd
			tmpRec(CurrentPage,index + 1,2) = tbldesc				
	   Case "del"
			tmpRec(CurrentPage,index + 1,0) = "del" 
	
End Select
Response.Write func & "<p>"
Response.Write tblcd & "-2" & "<p>"
Response.Write tbldesc & "-3" & "<p>"
Session("YDBSB0001EMP") = tmpRec
%>
</BODY>
</HTML>
