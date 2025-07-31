<%@ Language=VBScript %>
<%
Response.Expires = 0	
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<%
func = request("func")
muser = request("muser") 
username = request("username")
password = request("password")
index = request("index")
tmpRec = Session("YSBHE0501")
CurrentPage = request("CurrentPage")

on error resume next

Select Case func
	   Case "username_change"
			tmpRec(CurrentPage,index + 1,0) = "update"
			tmpRec(CurrentPage,index + 1,1) = muser
			tmpRec(CurrentPage,index + 1,2) = username
			tmpRec(CurrentPage,index + 1,3) = password
				
	   Case "del"
			tmpRec(CurrentPage,index + 1,0) = "del" 
	
End Select
Response.Write tmpRec(CurrentPage,index + 1,0)
Session("YSBHE0501") = tmpRec
%>
</BODY>
</HTML>
