<%@ Language=VBScript %>
<%
Response.Expires = 0	
%>
<HTML>
<HEAD>
<meta http-equiv="refresh">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
<BODY>
<%
func = request("func")
'tblcd = request("tblcd")
tbldesc = request("tbldesc") 
username= request("username") 
pwd = request("pwd") 
index = request("index") 
tmpRec = Session("YSBHE0502")
CurrentPage = request("CurrentPage") 


'on error resume next

Select Case func
	   Case "datachg"			
			tmpRec(CurrentPage,index + 1,0) = "upd"
			tmpRec(CurrentPage,index + 1,2) = username
			tmpRec(CurrentPage,index + 1,3) = tbldesc
			tmpRec(CurrentPage,index + 1,5) = pwd
		Case "del"			
			tmpRec(CurrentPage,index + 1,0) = "del"
		Case "no"			
			tmpRec(CurrentPage,index + 1,0) = "no"		
	
End Select
Response.Write "index = " & index &"<BR>"
Response.Write "0-" & tmpRec(CurrentPage,index + 1,0) &"<BR>"
Response.Write "3-" & tmpRec(CurrentPage,index + 1,3) &"<BR>"
Session("YSBHE0502") = tmpRec
%>
</BODY>
</HTML>
