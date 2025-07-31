<%@ Language=VBScript codepage=65001%>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<%

func = request("func")
tblcd = request("tblcd")
tbldesc = request("tbldesc")
index = request("index")
CurrentPage = request("CurrentPage")
memo2 = request("memo2")
tscode = request("tscode")
Response.Write func & "<p>"
Response.Write tblcd & "-2" & "<p>"
Response.Write tbldesc & "-3" & "<p>"
Response.Write currentPage & "<p>"
Response.Write index & "-index <p>"

tmpRec = Session("YTBAE0101EMP")

Select Case func
	   Case "tblcd_change"
			tmpRec(CurrentPage,index + 1,0) = "update"
			tmpRec(CurrentPage,index + 1,1) = tblcd
			tmpRec(CurrentPage,index + 1,2) = tbldesc				
			tmpRec(CurrentPage,index + 1,3) = memo2				
			tmpRec(CurrentPage,index + 1,5) = tscode		
	   Case "del"
			tmpRec(CurrentPage,index + 1,0) = "del" 
	
End Select
Response.Write func & "<p>"
Response.Write tblcd & "-2" & "<p>"
Response.Write tbldesc & "-3" & "<p>"
Session("YTBAE0101EMP") = tmpRec
%>
 
