<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" --> 
<%
SELF = "EMPBASIC" 

FUNC = request("func") 
code = request("code") 
code1= request("code1") 
code2= request("code2") 
code3= request("code3") 

code4= request("code4") 
code5= request("code5") 
code6= request("code6") 
code7= request("code7") 
code8= request("code8") 

dat1 = request("dat1")
dat2 = request("dat2")

response.write "dat1 = " & dat1 &"<BR>"
response.write "dat2 = " & dat2 &"<BR>"
index=request("index")  
CurrentPage = request("CurrentPage")  

tmpRec = Session("YEBE0104B")  


response.write "index=" & index &"<BR>"

'Set conn = GetSQLServerConnection()	 
%>
<html>
<head>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh"> 
</head>
<%
select case FUNC 
	case "yes"		
		tmpRec(CurrentPage,index + 1,0) = "Y"		 	
		tmpRec(CurrentPage,index + 1,6) = dat1 
		tmpRec(CurrentPage,index + 1,8) = dat2 
response.write CurrentPage &"-"&index &"-"&"0-" & tmpRec(CurrentPage,index + 1,0) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"1-" & tmpRec(CurrentPage,index + 1,1) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"2-" & tmpRec(CurrentPage,index + 1,2) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"3-" & tmpRec(CurrentPage,index + 1,3) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"4-" & tmpRec(CurrentPage,index + 1,4) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"5-" & tmpRec(CurrentPage,index + 1,5) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"6-" & tmpRec(CurrentPage,index + 1,6) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"7-" & tmpRec(CurrentPage,index + 1,7) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"8-" & tmpRec(CurrentPage,index + 1,8) &"<BR>" 		
	case "no"		
		tmpRec(CurrentPage,index + 1,0) = "no"
	case "datachg"		
		'tmpRec(CurrentPage,index + 1,0) = "no"		 
		tmpRec(CurrentPage,index + 1,6) = dat1 
		tmpRec(CurrentPage,index + 1,8) = dat2 
response.write CurrentPage &"-"&index &"-"&"0-" & tmpRec(CurrentPage,index + 1,0) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"1-" & tmpRec(CurrentPage,index + 1,1) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"2-" & tmpRec(CurrentPage,index + 1,2) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"3-" & tmpRec(CurrentPage,index + 1,3) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"4-" & tmpRec(CurrentPage,index + 1,4) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"5-" & tmpRec(CurrentPage,index + 1,5) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"6-" & tmpRec(CurrentPage,index + 1,6) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"7-" & tmpRec(CurrentPage,index + 1,7) &"<BR>"
response.write CurrentPage &"-"&index &"-"&"8-" & tmpRec(CurrentPage,index + 1,8) &"<BR>" 
		
end  select   		

Session("YEBE0104B") = tmpRec
%>

</html>

