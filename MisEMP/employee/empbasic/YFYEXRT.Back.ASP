<%@ Language=VBScript %>
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
yyyymm = request("yyyymm")

code = request("code")
if request("exrt") = "" then
   exrt = 0
else   
   exrt = request("exrt")
end if
index = request("index")
CurrentPage = request("CurrentPage")

tmpRec = Session("YSBAE1701")

Select Case func
	   Case "TaxCode_change"
			tmpRec(CurrentPage,index + 1,0) = "update"
			tmpRec(CurrentPage,index + 1,1) = yyyymm
			tmpRec(CurrentPage,index + 1,2) = code
			tmpRec(CurrentPage,index + 1,3) = exrt
				
	   Case "del"
			tmpRec(CurrentPage,index + 1,0) = "del" 
	
End Select

Session("YSBAE1701") = tmpRec


%>
</BODY>
</HTML>
