<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "YEDE02"  
ftype = request("func") 
code = request("code") 
index=request("index")  
CurrentPage = request("CurrentPage") 

CODESTR01 = REQUEST("code01")
CODESTR02 = REQUEST("CODE02")
CODESTR03 = REQUEST("CODE03")
CODESTR04 = REQUEST("CODE04")
CODESTR05 = REQUEST("CODE05")
CODESTR06 = REQUEST("CODE06")
CODESTR07 = REQUEST("CODE07")
CODESTR08 = REQUEST("CODE08")

tmpRec = Session("YEDE02B") 
response.write "index=" & index &"<BR>"
response.write "ftype=" & ftype &"<BR>"
Set conn = GetSQLServerConnection()	 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh"> 
</head>
<%
select case ftype 
	case "upd"				
  		tmpRec(CurrentPage,index + 1,0) ="*"
  		tmpRec(CurrentPage,index + 1,2) = CODESTR01
	  	tmpRec(CurrentPage,index + 1,3) = CODESTR02
	  	tmpRec(CurrentPage,index + 1,7) = CODESTR08
  		tmpRec(CurrentPage,index + 1,9) = CODESTR03
  		tmpRec(CurrentPage,index + 1,10) = CODESTR04
  		tmpRec(CurrentPage,index + 1,12) = CODESTR05
  		tmpRec(CurrentPage,index + 1,16) = CODESTR07
  		tmpRec(CurrentPage,index + 1,25) = CODESTR06
end  select   		
response.write  index &"<BR>"
response.write  "0-"&tmpRec(CurrentPage,index + 1,0)  &"<BR>"
response.write  "2-"&tmpRec(CurrentPage,index + 1,2)  &"<BR>"
response.write  "3-"&tmpRec(CurrentPage,index + 1,3)  &"<BR>"
response.write  "7-"&tmpRec(CurrentPage,index + 1,7)  &"<BR>"
response.write  "9-"&tmpRec(CurrentPage,index + 1,9)  &"<BR>"
response.write  "10-"&tmpRec(CurrentPage,index + 1,10)  &"<BR>"
response.write  "12-"&tmpRec(CurrentPage,index + 1,12)  &"<BR>"
response.write  "16-"&tmpRec(CurrentPage,index + 1,16)  &"<BR>"
response.write  "25-"&tmpRec(CurrentPage,index + 1,25)  &"<BR>"
Session("YEDE02B") = tmpRec
%>
</html>
